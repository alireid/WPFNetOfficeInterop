# WPFNetOfficeInterop

Please also see:
Demonstration of custom VSTO .NET C# AddIn 
Link: [VSTOWordRibbon](https://github.com/alireid/VSTOWordRibbon "VSTOWordRibbon")

------------

- Demonstration of MS Office interop using .NET WPF C#.
- This demo only uses the MS Office Interop capabilities within the .NET framework. It is worth noting that ClosedXML is another method of outputting office based documents using XML (OpenXML API). As opposed to interop, when using ClosedXML there is no necessity for MS Office applications to be installed on client/server platforms.
- This is a WPF application using the Model–view–viewmodel pattern. In the case here ~/Model/User.cs is the model, ~/ViewModel/UserViewModel.cs is the ViewModel which is a container for objects and view related functions that are ultimately passed to the view ~/View/MainPage.xaml
- Files that are output from the application are saved to C:\temp. If the directory does not exist it will be automatically created.
- Visual Studio 2019 solution.