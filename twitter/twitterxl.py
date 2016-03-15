from pyxll import RTD, xl_func, xl_on_close
from tweepy.streaming import StreamListener
from tweepy import OAuthHandler
from tweepy import Stream
import threading
import itertools
import logging
import json

_log = logging.getLogger(__name__)

# User credentials to access Twitter API
access_token = "YOUR ACCESS TOKEN"
access_token_secret = "YOUR ACCESS TOKEN SECRET"
consumer_key = "YOUR CONSUMER KEY"
consumer_secret = "YOUR CONSUMER KEY SECRET"


class TwitterListener(StreamListener):
    """tweepy.StreamListener that notifies multiple subscribers when
    new tweets are received and keeps a buffer of the last 100 tweets
    received.
    """
    __listeners = {}  # class level cache of listeners, keyed by phrases
    __lock = threading.RLock()
    __max_size = 100

    @classmethod
    def get_listener(cls, phrases, subscriber):
        """Fetch an ExcelListener listening to a set of phrases and subscribe to it"""
        with cls.__lock:
            # get the listener from the cache or create a new one
            phrases = frozenset(map(str, phrases))
            listener = cls.__listeners.get(phrases, None)
            if listener is None:
                listener = cls(phrases)
                cls.__listeners[phrases] = listener

            # add the subscription and return
            listener.subscribe(subscriber)
            return listener

    def __init__(self, phrases):
        """Use static method 'get_listener' instead of constructing directly"""
        _log.info("Creating listener for [%s]" % ", ".join(phrases))
        self.__phrases = phrases
        self.__subscriptions = set()
        self.__tweets = [None] * self.__max_size

        # listen for tweets in a background thread using the 'async' keyword
        auth = OAuthHandler(consumer_key, consumer_secret)
        auth.set_access_token(access_token, access_token_secret)
        self.__stream = Stream(auth, listener=self)
        self.__stream.filter(track=phrases, async=True)
        self.__connected = True

    @property
    def tweets(self):
        return list(self.__tweets)

    def subscribe(self, subscriber):
        """Add a subscriber that will be notified when new tweets are received"""
        with self.__lock:
            self.__subscriptions.add(subscriber)

    def unsubscribe(self, subscriber):
        """Remove subscriber added previously.
        When there are no more subscribers the listener is stopped.
        """
        with self.__lock:
            self.__subscriptions.remove(subscriber)
            if not self.__subscriptions:
                self.disconnect()

    def disconnect(self):
        """Disconnect from the twitter stream and remove from the cache of listeners."""
        with self.__lock:
            if self.__connected:
                _log.info("Disconnecting twitter stream for [%s]" % ", ".join(self.__phrases))
                self.__listeners.pop(self.__phrases)
                self.__stream.disconnect()
                self.__connected = False

    @classmethod
    def disconnect_all(cls):
        """Disconnect all listeners."""
        with cls.__lock:
            for listener in list(cls.__listeners.values()):
                listener.disconnect()

    def on_data(self, data):
        data = json.loads(data)
        with self.__lock:
            self.__tweets.insert(0, data)
            self.__tweets = self.__tweets[:self.__max_size]
            for subscriber in self.__subscriptions:
                try:
                    subscriber.on_data(data)
                except:
                    _log.error("Error calling subscriber", exc_info=True)
        return True

    def on_error(self, status):
        with self.__lock:
            for subscriber in self.__subscriptions:
                try:
                    subscriber.on_error(status)
                except:
                    _log.error("Error calling subscriber", exc_info=True)


# Make sure all the listeners get disconnected when Excel is closing
@xl_on_close
def disconnect_all_listeners():
    TwitterListener.disconnect_all()


class TwitterRTD(RTD):
    """Twitter RTD class that notifies Excel whenever a new tweet is received."""

    def __init__(self, phrases, row=0, key="text"):
        super(TwitterRTD, self).__init__(value="Waiting for tweets...")
        self.__listener = TwitterListener.get_listener(phrases, self)
        self.__row = row
        self.__key = key

    def disconnect(self):
        self.__listener.unsubscribe(self)

    def on_error(self, status):
        self.value = "#ERROR %s" % status

    def on_data(self, data):
        # if there are no tweets for this row return an empty string
        tweets = self.__listener.tweets
        if len(tweets) < self.__row or not tweets[self.__row]:
            self.value = ""
            return

        # get the value from the tweets
        value = tweets[self.__row]
        for key in self.__key.split("/"):
            if not isinstance(value, dict):
                value = ""
                break
            value = value.get(key, {})

        # set the value back in Excel
        self.value = str(value)



@xl_func("string[] phrases, int row, string key: rtd")
def twitter_listen(phrases, row=0, key="text"):
    """Listen for tweets containing certain phrases"""
    # flatten the 2d list of lists into a single list of phrases
    phrases = [str(x) for x in itertools.chain(*phrases) if x]
    assert len(phrases) > 0, "At least one phrase is required"

    # return our TwitterRTD object that will update when a tweet is received
    return TwitterRTD(phrases, row, key)


if __name__ == '__main__':
    import time
    logging.basicConfig(level=logging.INFO)

    class TestSubscriber(object):
        """simple subscriber that just prints tweets as they arrive"""

        def on_error(self, status):
            print("Error: %s" % status)

        def on_data(self, data):
            print(data.get("text"))

    subscriber = TestSubscriber()
    listener = TwitterListener.get_listener(["python", "excel", "pyxll"], subscriber)

    # listen for 60 seconds then stop
    time.sleep(60)
    listener.unsubscribe(subscriber)
