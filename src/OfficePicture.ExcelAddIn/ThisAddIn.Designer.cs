using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePicture.ExcelAddIn;

[Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObject(0)]
public sealed partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
{
    internal Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes = null!;
    internal Excel.Application Application = null!;
    public ThisAddIn(Microsoft.Office.Tools.Excel.ApplicationFactory factory, IServiceProvider serviceProvider) : base(factory, serviceProvider, "AddIn", "ThisAddIn") => Globals.Factory = factory;
    protected override void Initialize() { base.Initialize(); Application = GetHostItem<Excel.Application>(typeof(Excel.Application), "Application"); Globals.ThisAddIn = this; System.Windows.Forms.Application.EnableVisualStyles(); CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this); }
    protected override void FinishInitialization() { InternalStartup(); OnStartup(); }
    protected override void OnShutdown() { CustomTaskPanes.Dispose(); base.OnShutdown(); }
    private void InternalStartup() { Startup += ThisAddIn_Startup; Shutdown += ThisAddIn_Shutdown; }
    private void ThisAddIn_Startup(object sender, EventArgs e) { }
    private void ThisAddIn_Shutdown(object sender, EventArgs e) { }
}
internal static class Globals { internal static ThisAddIn ThisAddIn = null!; internal static Microsoft.Office.Tools.Excel.ApplicationFactory Factory = null!; }
