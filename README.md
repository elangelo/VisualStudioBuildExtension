# VisualStudioBuildExtension
Run a powershell script before and after the solution build

## How to use
Place a `SolutionName.PreBuild.ps1` and/or `SolutionName.PostBuild.ps1` into the same directory of your solution (`SolutionName.sln`). These scripts will be ran before or after the build of your solution completes.


> &#9888; Make sure these scripts are already there when you open your solution. They are cached for the lifetime of your Visual Studio session. If you need to change the scripts you will need to close and open Visual Studio again.

To see what is happening an extra output pane is opened in Visual Studio so you can follow what is happening with your scripts.

If your scripts needs Administrator access you need to run your Visual Studio session as administrator.