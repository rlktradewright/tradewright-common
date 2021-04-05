# tradewright-common
A set of utility libraries for use with COM-capable development environments
(VB6, VBA, .Net etc). These have been developed over the years since 2003 to provide
a number of useful features that are not possible 'out of the box' with VB6. They include:

* a powerful logging facility, based heavily on the Java Logging Framework
* sophisticated mechanisms for using configuration files (typically based on XML, though
other providers could be developed)
* clocks that operate in any timezone, and that may use simulated time for re-running
historical scenarios. Also functions for converting times between any timezones
* timers for accurate single or periodic notifications
* high-resolution elapsed time measurement
* futures
* parameter collections (name/value pairs)
* sorted dictionary
* state transition engine
* weak references to avoid the circular reference problem
* error handling with stack traces at point of failure
* cooperative multi-tasking to allow long-running activities on the main thread to be
interleaved, and without blocking the user interface
* a deferred action mechanism that enables actions to be postponed to a specified later time
* procedure call tracing 
* enumerators
* a subclassing framework
* a 'proper' graphics library (including gradient fills, advanced typography etc)
* a number of UI controls that provide advantages over the Microsoft equivalents

For information on building this project, see
[How to Build Tradewright Common](HowToBuildTradewrightCommon.md).

If you make changes that you want to contribute to the official version, create a pull
request and I'll evaluate them.

There is a lot more to be said about working with this code, because of the niceties of
such things as binary compatibility, but I won't say it now because I doubt anyone will
ever need it, but if you do, feel free to contact me or raise an issue.

