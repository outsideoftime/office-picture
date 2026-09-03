using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePicture.PowerPointAddIn;

[Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObject(0)]
public sealed partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
{
    internal Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes = null!;
    internal PowerPoint.Application Application = null!;
    public ThisAddIn(Microsoft.Office.Tools.Factory factory, IServiceProvider serviceProvider) : base(factory, serviceProvider, "AddIn", "ThisAddIn") => Globals.Factory = factory;
    protected override void Initialize() { base.Initialize(); Application = GetHostItem<PowerPoint.Application>(typeof(PowerPoint.Application), "Application"); Globals.ThisAddIn = this; System.Windows.Forms.Application.EnableVisualStyles(); CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this); }
    protected override void FinishInitialization() { InternalStartup(); OnStartup(); }
    protected override void OnShutdown() { CustomTaskPanes.Dispose(); base.OnShutdown(); }
    protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() => new PowerPointRibbon(this);
    private void InternalStartup() { Startup += ThisAddIn_Startup; Shutdown += ThisAddIn_Shutdown; }
}
internal static class Globals { internal static ThisAddIn ThisAddIn = null!; internal static Microsoft.Office.Tools.Factory Factory = null!; }
