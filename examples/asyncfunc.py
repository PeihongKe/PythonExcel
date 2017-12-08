"""
PyXLL Examples: Async function

Starting with Excel 2010 worksheet functions can
be registered as asynchronous.

This can be used for querying results from a server
asynchronously to improve the worksheet calculation
performance.
"""

from pyxll import xl_func, xl_version, xlAsyncReturn

#
# this example uses urllib2 to perform an asynchronous http
# request and return the data to Excel.
#
import urllib.request, urllib.error, urllib.parse
import threading

#
# Async functions are only supported from Excel 2010
#
if xl_version() >= 14:

    @xl_func("string, async_handle: void")
    def yahoo_stock_price(symbol, handle):
        """returns the last price for a symbol from Yahoo Finance"""

        def thread_func(symbol, async_handle):
            result = None
            try:
                # get the price using an http request (f=l1 means get the last price)
                url = "http://download.finance.yahoo.com/d/quotes.csv?s=%s&f=l1" % symbol
                data = urllib.request.urlopen(url).read()
                
                # the returned data is in csv format, but only has one row and one column
                result = float(data.strip())
            except Exception as e:
                result = e

            # return the result to Excel
            xlAsyncReturn(handle, result)
 
        # do the request in a new thread
        thread = threading.Thread(target=thread_func, args=(symbol, handle))
        thread.start()

        # done, no need to return a value as it is done by the response handler
        return

else:

    @xl_func("string: string")
    def yahoo_stock_price(symbol):
        """not supported in this version of Excel"""
        return "async functions are not supported in Excel %s" % xl_version()