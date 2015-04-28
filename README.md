# tradewright-common
A set of utility libraries for use with COM-capable development environments (VB6, VBA, .Net etc). These have been developed over the years since 2003 to provide a number of useful features that are not possible 'out of the box' with VB6. They include:

* a powerful logging facility, based heavily on the Java Logging Framework
* sophisticated mechanisms for using configuration files (typically based on XML, though other providers could be developed)
* clocks that operate in any timezone, and that may use simulated time for re-running historical scenarios. Also functions for converting times between any timezones
* timers for accurate single or periodic notifications
* high-resolution elapsed time measurement
* futures
* parameter collections (name/value pairs)
* sorted dictionary
* state transition engine
* weak references to avoid the circular reference problem
* error handling with stack traces at point of failure
* cooperative multi-tasking to allow long-running activities on the main thread to be interleaved, and without blocking the user interface
* a deferred action mechanism that enables actions to be postponed to a specified later time
* procedure call tracing 
* enumerators
* a subclassing framework
* a 'proper' graphics library (including gradient fills, advanced typography etc)
* a number of UI controls that provide advantages over the Microsoft equivalents

During development of this software until February 2015, Visual SourceSafe (VSS) was used for version control. The advantages of git in general and GitHub in particular have made it seem worthile to move the project here, even though there is no git integration in Visual Basic 6. I have made no attempt to carry the history from VSS over to GitHub - as the sole developer, the history is unlikely to be of any interest to anyone else, and I still have it in the VSS repository.

Having this on GitHub will also ease the process of porting relevant parts of it to .Net: this process has been underway for some time, but I intend to push it out to GitHub in the not-too-distant future.

I don't expect there to be much change here in future. The software is pretty robust, and of course the VB6 development platform is pretty much a dead duck (at least as far as Microsoft is concerned: though I have to say it's one of the most lively dead ducks I've ever come across - it just won't lie down and die properly!). I also don't really expect anyone else to have the slightest interest in it!

If by chance you do want to modify the software, here's what to do:

* Clone the repository to your computer. Note that you'll need Visual Basic 6 to be able to compile the code.

* Set up the following environment variables:
  
  `TW-PROJECTS-DRIVE` - set this to the drive containing your repository clone, eg C:

  `TW-PROJECTS-PATH` - set this to the path to your repository folder, eg \Projects\tradewright-common

* Now run the `registerTradeWrightCommon.bat` file in the `Build` folder: note you should run this as Administrator to avoid a series of privilege elevation prompts (you'll still get one of course!)

* You should now find that the sample .exe files in the Bin folder run nicely.

If you make changes that you want to contribute to the official version, create a pull request and I'll evaluate them.

There is a lot more to be said about working with this code, because of the niceties of such things as binary compatibility, but I won't say it now because I doubt anyone will ever need it, but if you do, feel free to contact me or raise an issue.

