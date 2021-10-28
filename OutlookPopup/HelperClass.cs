using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OutlookPopup
{
    class OfficeWin32Window
    {
        [DllImport("user32")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        #region IWin32Window Members

        IntPtr _windowHandle = IntPtr.Zero;
        public IntPtr Handle
        {
            get { return _windowHandle; }
        }

        #endregion


        public OfficeWin32Window(object windowObject)
        {
            string caption = windowObject.GetType().InvokeMember("Caption", System.Reflection.BindingFlags.GetProperty, null, windowObject, null).ToString();

            // try to get the HWND ptr from the windowObject / could be an Inspector window or an explorer window
            _windowHandle = FindWindow("rctrl_renwnd32\0", caption);
        }
    }
}

    