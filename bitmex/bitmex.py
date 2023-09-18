"""
Get BitMEX prices into Excel using websockets and PyXLL RTD functions.

Requires:
- Python 3.5 or higher
- websockets

>> pip install websockets

See https://pyxll.medium.com/live-streaming-crypto-prices-in-excel-aaa41628bc53
"""
from pyxll import xl_func, RTD, get_event_loop
import websockets
import asyncio
import json


class BitMex:
    """Class to manage subscriptions to instrument prices."""

    URI = "wss://www.bitmex.com/realtime"

    def __init__(self, loop=None):
        self.__websocket = None
        self.__running = False
        self.__running_task = None
        self.__subscriptions = {}
        self.__data = {}
        self.__lock = asyncio.Lock()

    async def __connect(self):
        # Connect to the websocket API and start the __run coroutine
        self.__running = True
        self.__websocket = await websockets.connect(self.URI)
        self.__connecting_task = None
        self.__running_task = asyncio.create_task(self.__run())

    async def __disconnect(self):
        # Close the websocket and wait for __run to complete
        self.__running = False
        await self.__websocket.close()
        self.__websocket = None
        await self.__running_task

    async def __run(self):
        # Read from the websocket until disconnected
        while self.__running:
            msg = await self.__websocket.recv()
            await self.__process_message(json.loads(msg))

    async def __process_message(self, msg):
        if msg.get("table", None) == "instrument":
            # Extract the data from the message, update our data dictionary and notify subscribers
            for data in msg.get("data", []):
                symbol = data["symbol"]
                timestamp = data["symbol"]

                # Update the latest values in our data dictionary and notify any subscribers
                tasks = []
                subscribers = self.__subscriptions.get(symbol, {})
                latest = self.__data.setdefault(symbol, {})
                for field, value in data.items():
                    latest[field] = (value, timestamp)

                    # Notify the subscribers with the updated field
                    for subscriber in subscribers.get(field, []):
                        tasks.append(subscriber(symbol, field, value, timestamp))

                # await all the tasks from the subscribers
                if tasks:
                    await asyncio.wait(tasks)

    async def subscribe(self, symbol, field, callback):
        """Subscribe to updates for a specific symbol and field.

        The callback will be called as 'await callback(symbol, field, value, timestamp)'
        whenever an update is received.
        """
        async with self.__lock:
            # Connect the websocket if necessary
            if self.__websocket is None:
                await self.__connect()

            # Send the subscribe message if we're not already subscribed
            if symbol not in self.__subscriptions:
                msg = {"op": "subscribe", "args": [f"instrument:{symbol}"]}
                await self.__websocket.send(json.dumps(msg))

            # Add the subscriber to the dict of subscriptions
            self.__subscriptions.setdefault(symbol, {}).setdefault(field, []).append(callback)

            # Call the callback with the latest data
            data = self.__data.get(symbol, {})
            if field in data:
                (value, timestamp) = data[field]
                await callback(symbol, field, value, timestamp)

    async def unsubscribe(self, symbol, field, callback):
        async with self.__lock:
            # Remove the subscriber from the list of subscriptions
            self.__subscriptions[symbol][field].remove(callback)
            if not self.__subscriptions[symbol][field]:
                del self.__subscriptions[symbol][field]

            # Unsubscribe if we no longer have any subscriptions for this instrument
            if not self.__subscriptions[symbol]:
                msg = {"op": "unsubscribe", "args": [f"instrument:{symbol}"]}
                await self.__websocket.send(json.dumps(msg))
                del self.__subscriptions[symbol]
                self.__data.pop(symbol, None)

            # Disconnect if we no longer have any subscriptions
            if not self.__subscriptions:
                async with self.__lock:
                    await self.__disconnect()


class BitMexRTD(RTD):
    """RTD class for subscribing to BitMEX prices using the
    BitMex class above.
    """

    # Use a single BitMex object for all RTD functions
    _bitmex = BitMex(get_event_loop())

    def __init__(self, symbol, field):
        super().__init__(value="Waiting...")
        self.__symbol = symbol
        self.__field = field

    async def connect(self):
        # Subscribe to BitMix updates when Excel connects to the RTD object
        await self._bitmex.subscribe(self.__symbol, self.__field, self.__update)

    async def disconnect(self):
        # Unsubscribe to BitMix updates when Excel disconnects from the RTD object
        await self._bitmex.unsubscribe(self.__symbol, self.__field, self.__update)

    async def __update(self, symbol, field, value, timestamp):
        # Update the value in Excel
        self.value = value


@xl_func("string symbol, string field: rtd", recalc_on_open=True)
def bitmex_rtd(symbol, field="lastPrice"):
    """Subscribe to BitMEX prices for a given symbol."""
    return BitMexRTD(symbol, field)


if __name__ == "__main__":

    async def main():
        # This is the callback that will be called whenever there's an update
        async def callback(symbol, field, value, timestamp):
            print((symbol, field, value, timestamp))

        bm = BitMex()

        await bm.subscribe("XBTUSD", "lastPrice", callback)

        await asyncio.sleep(60)

        await bm.unsubscribe("XBTUSD", "lastPrice", callback)

        print("DONE!")

    # Run the 'main' function in an asyncio event loop
    loop = asyncio.get_event_loop()
    loop.create_task(main())
    loop.run_forever()
