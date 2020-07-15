"""
Excel functions "publish" and "subscribe" for passing
values between workbooks.

An RTD function is used for subscribe so that published
values are updated in read time.
"""
from pyxll import xl_func, RTD
import pubsub


class SubscriberRTD(RTD):
    """RTD class that updates whenever a value is published
    on a given topic."""

    def __init__(self, topic):
        super().__init__(value="Waiting...")
        self.__topic = topic

    def connect(self):
        pubsub.subscribe(self.__topic, self.__callback)

    def disconnect(self):
        pubsub.unsubscribe(self.__topic, self.__callback)

    def __callback(self, value):
        self.value = value


@xl_func
def publish(topic, value):
    """Publish a value on a topic to be picked up by another sheet"""
    pubsub.publish(topic, value)
    return f"[Published to {topic}]"


@xl_func("str: rtd")
def subscribe(topic):
    """Get a value that has been published on another sheet"""
    return SubscriberRTD(topic)


#
# Example application:
#
# One sheet constructs a swap curve object and publishes it.
# Other sheets can subscribe to the topic the curve object
# is published on to get the curve, without having to
# reference the sheet directly or use INDEX.
#

class DummySwapCurve:
    def __init__(self, name):
        self.__name = name


@xl_func(recalc_on_open=True)
def create_swap_curve(name, args):
    """Creates a new swap curve and returns it to Excel"""
    return DummySwapCurve(name)
