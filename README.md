# WPFNetOfficeInterop

- Demonstration of MS Office interop using .NET WPF C#.
- This demo only uses the MS Office Interop capabilities within the .NET framework. It is worth noting that ClosedXML is a another method of outputting office based documents and uses XML (OpenXML API). Unlike Interop, when using ClosedXML there is no necessity for MS Office applications to be installed on client/server platforms.
- This is a WPF application using the Model–view–viewmodel pattern. In the case here ~/Model/User.cs is the model, ~/ViewModel/UserViewModel.cs is the ViewModel which is a container for objects that are ultimately passed to the view ~/View/MainPage.xaml
- Files that are output from the application are saved to C:\temp. The directory is automatically created if it doesn't exist.
- Visual Studio 2019 solution
