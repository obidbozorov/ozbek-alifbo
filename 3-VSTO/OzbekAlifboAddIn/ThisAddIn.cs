using System;
using Office = Microsoft.Office.Core;

namespace OzbekAlifboAddIn
{
    public partial class ThisAddIn
    {
        internal LiveTyping Live { get; private set; }
        internal Hotkeys Keys { get; private set; }
        internal AlifboRibbon Ribbon { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Live = new LiveTyping(Application);
            if (Settings.LiveTyping) Live.Enable();

            // Tezkor tugmalar: Alt + harf
            Keys = new Hotkeys();
            var r = Ribbon ?? (Ribbon = new AlifboRibbon());
            Keys.Register(System.Windows.Forms.Keys.A, () => r.ConvertAuto(ConvertMode.Latin));
            Keys.Register(System.Windows.Forms.Keys.E, () => r.ConvertAuto(ConvertMode.NewToOld));
            Keys.Register(System.Windows.Forms.Keys.K, () => r.ConvertAuto(ConvertMode.Cyrillic));
            Keys.Register(System.Windows.Forms.Keys.C, () => r.ConvertAuto(ConvertMode.LatinToCyrillic));
            Keys.Register(System.Windows.Forms.Keys.Z, r.ToggleLive);
            Keys.Register(System.Windows.Forms.Keys.I, r.AddExceptionAuto);
            Keys.Register(System.Windows.Forms.Keys.O, r.OpenExceptions);
            if (Settings.Hotkeys) Keys.Enable();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Live?.Dispose();
            Keys?.Dispose();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return Ribbon ?? (Ribbon = new AlifboRibbon());
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
