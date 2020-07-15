
# A Non-Volatile INDIRECT Alternative in Excel using the Pub/Sub Pattern

This code accompanies the Medium post
[A Non-Volatile INDIRECT Alternative in Excel using the Pub/Sub Pattern](https://medium.com/@tony_86605/a-non-volatile-indirect-alternative-in-excel-using-the-pub-sub-pattern-15cea21272a3)

Messages are posted to a message broker from one sheet and subscribed to in another
to pass values between Excel workbooks.

Updates to values are published and updated in the subscribing sheets(s) in
real time using an RTD function.

Unlike the Excel INDEX function this does not use a volatile function and
so doesn't suffer from the same potential performance problems that sheets
using INDEX do.
