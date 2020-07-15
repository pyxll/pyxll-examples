"""
A simple in-process message broker to decouple
components in a single Python application.
"""
import threading
import logging

_log = logging.getLogger(__name__)


class MessageBroker:
    """The message broker waits for topics to be published and then
    notifies any subscribers of the published topics.

    When a new subscriber subscribes they will be immediately sent the latest
    value if there is one.

    Only the latest value is used and so publishing multiple messages
    on the same topic may result in the subscriber only seeing the
    latest value. This is the intended behaviour.
    """

    def __init__(self):
        # Dict of current values for each topic.
        self.__latest_values = {}

        # Lists of callbacks for each topic subscribed to
        self.__subscribers = {}

        # Prevent concurrent access to the above data structures
        self.__lock = threading.RLock()

    def subscribe(self, topic, callback):
        """Subscribes to a topic.
        The subscriber callback is called with the topic value whenever
        the a value is published to that topic, and with the initial
        value if there is one.
        """
        with self.__lock:
            # Add the subscriber to the list of subscribers
            subscribers = self.__subscribers.setdefault(topic, [])
            subscribers.append(callback)

            # Get the current value for the topic, if there is one and call the callback
            value = self.__latest_values.get(topic)
            if value is not None:
                callback(value)

    def unsubscribe(self, topic, callback):
        """Unsubscribe from a topic.
        The topic and callback must have previously be passed to 'subscribe'.
        """
        with self.__lock:
            # Remove the callback from the list of subscribers
            subscribers = self.__subscribers[topic]
            subscribers.remove(callback)

            if not subscribers:
                del self.__subscribers[topic]

    def publish(self, topic, value):
        """Publish a value to all subscribers of a topic.
        """
        with self.__lock:
            # Call the subscribers for the published topic
            if value is not None:
                for callback in self.__subscribers.get(topic, []):
                    try:
                        callback(value)
                    except:
                        _log.error(f"Failed to notify subscriber of {topic}", exc_info=True)

            # Update the latest value for any new subscribers
            self.__latest_values[topic] = value


# Use a single global message broker for all topics
_global_message_broker = MessageBroker()


def publish(topic, value):
    """Publishes a value on a topic for a subscriber to receive"""
    _global_message_broker.publish(topic, value)


def subscribe(topic, callback):
    """Subscribes to a topic.
    The subscriber callback is called with the topic value whenever
    the a value is published to that topic, and with the initial
    value if there is one.
    """
    _global_message_broker.subscribe(topic, callback)


def unsubscribe(topic, callback):
    """Unsubscribe from a topic.
    The topic and callback must have previously be passed to 'subscribe'.
    """
    _global_message_broker.unsubscribe(topic, callback)
