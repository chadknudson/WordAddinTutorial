using NetOffice.Tools;
using NetOffice.WordApi;
using NetOffice.WordApi.Tools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace WordAddinTutorial
{
    [ComVisible(true)]
    [Guid("0014e22c-983a-4c6b-92ec-e39ba1db206e")]
    [ProgId("WordAddinTutorial.MyAddin")]
    [COMAddin("MyAddin", "Addin description.", LoadBehavior.LoadAtStartup)]
    public class MyAddin : COMAddin
    {
        public MyAddin()
        {
            this.OnConnection += MyAddin_OnConnection;
        }

        private void MyAddin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            this.Application.DocumentOpenEvent += Application_DocumentOpenEvent;
        }

        private void Application_DocumentOpenEvent(Document doc)
        {
            using (doc)
            {
                // start working with the document
            }
        }
    }
}
