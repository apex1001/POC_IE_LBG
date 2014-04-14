/* Proof of concept for the LBG Browser helper object.
 * Highlights all phone numbers in an HTML page
 * 
 * Author: V. Vogelesang
 * 
 * Based on the BHO tutorials @
 * http://www.codeguru.com/csharp/.net/net_general/comcom/article.php/c19613/Build-a-Managed-BHO-and-Plug-into-the-Browser.htm
 * http://www.codeproject.com/Articles/19971/How-to-attach-to-Browser-Helper-Object-BHO-with-C
 * 
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SHDocVw;
using mshtml;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;
using System.Collections;
using System.Text.RegularExpressions;

namespace POC_LBG
{
    [
        ComVisible(true),
        Guid("81574025-b790-4efc-a145-f67ad0c6fa47"),
        ClassInterface(ClassInterfaceType.None)
    ]
        
    public class BHO_lbg : IObjectWithSite
    {
        private SHDocVw.WebBrowser browserWindow;
        private HTMLDocument document;
        private int parseCount;

        private static string BHOROOTKEYNAME = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";
        ArrayList forbiddenTags = new ArrayList { "script", "style", "img", "audio", "table", "time", "video" };
        String regex = "(0|\\+|\\(\\+|\\(0)[0-9- ()]{9,}";
        String url = "http://callsomenumber/?number=";

        /**
         * Invoked on completion of document loading
         */
        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            if (pDisp != this.browserWindow)
                return;

            parseCount = 0;
            HTMLDocument document = (HTMLDocument)browserWindow.Document;
            if (document != null)
            {

                // Add events for hooking into DHTML DOM events
                HTMLWindowEvents2_Event windowEvent = (document.parentWindow as HTMLWindowEvents2_Event);
                HTMLDocumentEvents2_Event docEvent = (document as HTMLDocumentEvents2_Event);
                try
                {
                    windowEvent.onload -= new HTMLWindowEvents2_onloadEventHandler(OnLoadHandler);
                    docEvent.onpropertychange -= new HTMLDocumentEvents2_onpropertychangeEventHandler(OnPropertyChangeHandler);
                    docEvent.onreadystatechange -= new HTMLDocumentEvents2_onreadystatechangeEventHandler(OnReadyStateChangeHandler);
                    browserWindow.DownloadBegin -= new DWebBrowserEvents2_DownloadBeginEventHandler(DownloadHandler);
                    docEvent.onmousemove -= new HTMLDocumentEvents2_onmousemoveEventHandler(MouseHandler);
                }
                catch { }

                windowEvent.onload += new HTMLWindowEvents2_onloadEventHandler(OnLoadHandler);
                docEvent.onpropertychange += new HTMLDocumentEvents2_onpropertychangeEventHandler(OnPropertyChangeHandler);
                docEvent.onreadystatechange += new HTMLDocumentEvents2_onreadystatechangeEventHandler(OnReadyStateChangeHandler);
                browserWindow.DownloadBegin += new DWebBrowserEvents2_DownloadBeginEventHandler(DownloadHandler);
            }
            HighLightPhoneNumbers();           
        }

        /**
         * Start the highlight process
         * 
         */
        public void HighLightPhoneNumbers() {

            // Prevent 2nd parse run
            if (parseCount == 1) return;            
            parseCount++;

            document = (HTMLDocument)browserWindow.Document;
            IHTMLElementCollection elements = document.body.all;            

            foreach (IHTMLElement el in elements)
            {
                if (!forbiddenTags.Contains(el.tagName.ToLower()))
                {
                    IHTMLDOMNode domNode = el as IHTMLDOMNode;
                    if (domNode.hasChildNodes())
                    {
                        IHTMLDOMChildrenCollection domNodeChildren = domNode.childNodes;
                        foreach (IHTMLDOMNode child in domNodeChildren)
                        {
                            if (child.nodeType == 3)
                            {
                                MatchCollection matches = Regex.Matches(child.nodeValue, regex);
                                if (matches.Count > 0)
                                {
                                    String newChildNodeValue = child.nodeValue;
                                    foreach (Match match in matches)
                                    {                                        
                                        String hlText = match.Value;
                                        newChildNodeValue = newChildNodeValue.Replace(hlText,
                                            "<a name=\"tel\" href=\"" + url + hlText + "\">" + hlText + "</a>");
                                    }
                                    IHTMLElement newChild = document.createElement("text");
                                    newChild.innerHTML = newChildNodeValue;
                                    child.replaceNode((IHTMLDOMNode)newChild);                                        
                                  
                                }
                            }
                        }
                    }
                }
            }            
            
            // Get all a elements wit phonenumber and add onclick evenhandler
            DispatcherClass dp = new DispatcherClass();            
            IHTMLElementCollection telElements = document.getElementsByName("tel");
            foreach (IHTMLElement el in telElements)
            {
                el.onclick = dp;
            }
            
            //System.Windows.Forms.MessageBox.Show("Elements: " + telElements.length);      
        }

        public void MouseHandler(IHTMLEventObj e)
        {
            HTMLDocument document = (HTMLDocument)browserWindow.Document;
            HTMLDocumentEvents2_Event docEvent = (document as HTMLDocumentEvents2_Event);
            docEvent.onmousemove -= new HTMLDocumentEvents2_onmousemoveEventHandler(MouseHandler);
            HighLightPhoneNumbers();
        }

        public void DownloadHandler()
        {
            HighLightPhoneNumbers();
        }

        public void OnReadyStateChangeHandler(IHTMLEventObj e)
        {
            HighLightPhoneNumbers();
        }

        public void OnPropertyChangeHandler(IHTMLEventObj e)
        {
            HTMLDocument document = (HTMLDocument)browserWindow.Document;
            HTMLDocumentEvents2_Event docEvent = (document as HTMLDocumentEvents2_Event);
            docEvent.onmousemove += new HTMLDocumentEvents2_onmousemoveEventHandler(MouseHandler);
            HighLightPhoneNumbers();
        }

        public void OnLoadHandler(IHTMLEventObj e)
        {
            HighLightPhoneNumbers();
        }

         /*
         * Set reference to the browser window on creation
         * 
         * @param browserwindow
         */
        public int SetSite(object site)
        {
            //this.webSite = site;
            browserWindow = (SHDocVw.WebBrowser)site;
            String windowName = browserWindow.FullName;
            
            // Browser window is in startup!
            // Register eventhandler to hook DocumentComplete() to own method                        
            if (browserWindow != null && windowName.ToUpper().EndsWith("IEXPLORE.EXE"))
            {                
                browserWindow.DocumentComplete +=
                    new DWebBrowserEvents2_DocumentCompleteEventHandler(
                        this.OnDocumentComplete); 
            }

            // Browser window is in shutdown!
            // Unregister eventhandler to prevent IE crashing
            // Also release browserWindow COM object reference
            else
            {
                browserWindow.DocumentComplete -=
                    new DWebBrowserEvents2_DocumentCompleteEventHandler(
                        this.OnDocumentComplete);
                Marshal.FinalReleaseComObject(browserWindow);
                browserWindow = null;                 
            }
            return 0;
        }

        /*
         * Get reference to last site from SetSite
         */
        public int GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            const int E_FAIL = -2147467259;
            if (browserWindow == null)
            {
                ppvSite = IntPtr.Zero;
                return E_FAIL;
            }            
            else
            {
                IntPtr punk = Marshal.GetIUnknownForObject(browserWindow);
                int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
                Marshal.Release(punk);
                return hr;
            }            
        }

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey BHORootKey = Registry.LocalMachine.OpenSubKey(BHOROOTKEYNAME, true);
            if (BHORootKey == null)
            {
                BHORootKey = Registry.LocalMachine.CreateSubKey(BHOROOTKEYNAME);
            }
            
            string thisGuid = type.GUID.ToString("B").ToUpper();
            RegistryKey LbgKey = BHORootKey.OpenSubKey(thisGuid);

            if (LbgKey == null)
            {
                LbgKey = BHORootKey.CreateSubKey(thisGuid);
            }
            LbgKey.SetValue("DummyKey", 1);
            LbgKey.Close();
            BHORootKey.Close();
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey BHORootKey = Registry.LocalMachine.OpenSubKey(BHOROOTKEYNAME, true);
            string thisGuid = type.GUID.ToString("B");

            if (BHORootKey != null)
                BHORootKey.DeleteSubKey(thisGuid, false);
        }

        
   
    }

    public class DispatcherClass
    {
        [DispId(0)]
        public void DefaultMethod()
        {
            MessageBox.Show("ja");
        }
    }

}
