# How to Build Tradewright Common

This document describes the steps that need to be taken to enable the TradeWright Common
components to be built.

It is quite a detailed procedure, but it only needs to be done once.

## 1. Prerequisites

Visual Basic 6 must be installed.

Visual Studio 2017 or later must be installed (earlier versions will probably be
fine, but haven't been tested). The free Community editions are fine and recommended.

## 2. Create the Project Environment

In order to get set up for building the TradeWright Common components, a number of Git
repositories have to be cloned from GitHub.

You can locate the cloned repositories anywhere you like on your computer.
However, it's recommended that you put them all in a common root folder, and for
consistency with these instructions this should be at `C:\Projects`.


## 3. Establish the Build Environment

1. Clone the vb6-build project from https://github.com/tradewright/vb6-build.
   This contains tools and scripts used in building TradeWright Common and other
   software.

2. Build the SetProjectComp tool using Visual Studio 2017 (or later) from the
   SetProjectComp.sln solution file in the `src\SetProjectComp folder`. Copy the
   generated SetProjectComp.exe file to the Tools folder of the vb6-build
   project. NB: this is a .Net program, so you can't build it with VB6.

3. Edit the system PATH variable to include the SCRIPTS and TOOLS folders in the
   vb6-build project.

4. Edit the system PATH variable to include:

   `C:\Program Files (x86)\Microsoft Visual Studio\VB98`

5. Clone the manifest-generator project from https://github.com/rlktradewright/manifest-generator.
   This project provides tools for generating assembly and program manifests
   for use with COM-free registration.

6. Build the manifest-generator solution using Visual Studio 2017 (or later)
   from the ManifestUtilities.sln solution file in the root folder. Copy the
   GenerateManifest.exe, ManifestUtilities.dll and Utils.dll files to the Tools
   folder of the vb6-build project. NB: this is a .Net solution, so you can't
   build it with VB6.


## 4. Setup the TradeWright Common Project

1. Clone the TradeWright Common  repository to your computer from GitHub. The link to
   the repository is:

   https://github.com/rlktradewright/tradewright-common

2. Install TradeWright Common using the .msi installer file from the latest Release on
   GitHub. By default, the files will be installed to:

   `C:\Program Files (x86)\TradeWright Software Systems\TradeWright Utilities Sample Apps vnnn`
   
   where nnn is the version number.

   Note that you only need to install TradeWright Common to give you initial versions to
   control binary compatibility during the first build process. Once you have
   successfully built TradeWright Common, you can uninstall it again.

3. Create a folder called Bin directly below the folder you cloned the
   repository to, for example:

   `C:\Projects\tradewright-common\Bin`

   This folder will be the root of a set of subfolders where the compiled
   TradeWright Common components components will be stored.

4. Create a folder called Compat directly below the folder you cloned the
   repository to, for example:

   `C:\Projects\tradewright-common\Compat`

   This folder will contain a parallel set of folders to those in the Bin
   folder, which will contain compiled TradeWright Common components components used for
   controlling binary compatibility.

5. Copy the  
   `TradeWright.Common` folder from the TradeWright Common installation folder tree
   to the Compat folder you just created.

8. Register the TradeWright common .dll's and .ocx's. We need to do this only so
   that other porjects that use the TradeWright Common user controls can
   be opened in Visual Basic 6 (note that Visual Basic 6 cannot use components
   via registration-free COM). For ease, we register the files in the
   TradeWright Common installation folders.

   * Open a Visual Studio developer command prompt as Administrator 
     to ensure that the regsvr32exe program is on the path). The easiest way to
	 do this on Windows 10 is to expand the Visual Studio 2017/2019 entry in the
	 Start menu, right-click the Developer Command Prompt item, and select
	 `More > Run as Administrator`.
	
   * Set the current directory to the TradeWright Common installation folder:
	
     `cd  "C:\Program Files\TradeWright Software Systems\TradeWright Utilities Sample Apps vnnn"`
	
   * Run the registerdlls.bat command file.

9. Create a user environment variable called `TW-PROJECTS-DRIVE`, and
   set it to the drive letter that contains your Projects folder (in this
   example set it to `C:`).

10. Create a user environment variable called `TW-PROJECTS-PATH`, and
   set it to the path to your clone of the TradeWright Common components repository (in this
   example set it to `\Projects\tradewright-common`).



## 5. Build TradeWright Common components for the First Time

1. You are now ready to do the first build of the TradeWright Common components project.

2. Start a Visual Studio developer command prompt as Administrator (as
   described earlier).

3. Set the current directory to the Build folder:

   `cd C:\projects\tradewright-common\Build`

4. Initiate the build:

   `makeall`

5. This will compile all the TradeWright Common components components, storing the resulting
   .exe's, dll's and ocx's into `C:\projects\tradewright-common\Build` and
   its various subfolders. It also generates COM interop components for all
   relevant items to enable them to be used seamlessly in .Net programs.

6. Once you have successfully built TradeWright Common, you can uninstall the version
   created by the .msi installer, as none of that installation is now in use.


