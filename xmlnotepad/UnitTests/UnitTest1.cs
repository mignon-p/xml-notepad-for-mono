using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Drawing;
using Accessibility;
using System.Diagnostics;
using System.Reflection;
using XmlNotepad;

// Here's a handy reference on SendKeys:
// http://msdn2.microsoft.com/en-us/library/system.windows.forms.sendkeys.aspx

namespace UnitTests {
    
    [TestClass]
    public class UnitTest1 : TestBase {
        string TestDir;
                
        public UnitTest1() {
            Uri baseUri = new Uri(this.GetType().Assembly.Location);
            Uri resolved = new Uri(baseUri, "..\\..\\..\\");
            TestDir = resolved.LocalPath;
            // Test that we can process updates and show available updates button.
            // Have to fix the location field to show the right thing.
            XmlDocument doc = new XmlDocument();
            doc.Load(TestDir + @"UnitTests\TestUpdates.xml");
            XmlElement e = (XmlElement)doc.SelectSingleNode("updates/application/location");
            string target = TestDir + @"XmlNotepad\bin\Debug\Updates.xml";
            e.InnerText = target;
            doc.Save(target);
        }

        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext) {
        //}
        
        //[ClassCleanup()]
        //public static void MyClassCleanup() {            
        //}
        
        [TestInitialize()]        
        public void MyTestInitialize() {
        }
        
        [TestCleanup()]
        public void MyTestCleanup() 
        {
            if (TestForm != null) {
                TestForm.Dispose();
            }
        }

        Window LaunchNotepad() {
            Window w = LaunchNotepad(null);
            TestForm.InvokeMenuItem("newToolStripMenuItem");
            Sleep(1000);
            return w;
        }

        Window LaunchNotepad(string filename) {
            Window wrapper = LaunchApp(Directory.GetCurrentDirectory() + @"\XmlNotepad.exe", filename);
            return wrapper;
        }


        SystemAccessible XmlTreeView {
            get {
                SystemAccessible xtv = this.TestForm.FindDescendant("xmlTreeView1");
                return xtv;
            }
        }

        SystemAccessible TreeView {
            get {
                SystemAccessible tv = this.TestForm.FindDescendant("TreeView");
                return tv;
            }
        }


        SystemAccessible NodeTextView {
            get {
                SystemAccessible ntv = this.TestForm.FindDescendant("NodeTextView");
                return ntv;
            }
        }

        SystemAccessible NodeTextViewCompletionSet {
            get {
                SystemAccessible ntv = this.NodeTextView;
                SystemAccessible cset = ntv.LastChild;
                if (cset.Name != "CompletionSet") {
                    throw new Exception("CompletionSet is missing");
                }
                if ((cset.State & AccessibleStates.Invisible) != 0) {
                    throw new Exception("CompletionSet is not visible");
                }
                return cset;
            }
        }
        
        [TestMethod]
        public void TestUndoRedo() {
            Trace.WriteLine("TestUndoRedo==========================================================");
            // Since this is the first test, we have to make sure we don't load some other user settings.
            string testFile = TestDir + "UnitTests\\test1.xml";
            this.LaunchNotepad(testFile);

            // test that we can cancel editing when we click New
            Sleep(500); 
            SendKeys.SendWait("^IRoot{ENTER}");
            Sleep(100); 
            TestForm.InvokeMenuItem("newToolStripMenuItem");
            Sleep(500);

            TestForm.SetWindowSize(800, 600);

            Stack<bool> hasChildren = new Stack<bool>();
            List<NodeInfo> nodes = new List<NodeInfo>();
            string dir = Directory.GetCurrentDirectory();
            XmlReader reader = XmlReader.Create(testFile);
            bool openElement = true;
            int commands = 0;
            bool readyForText = false;
            IntPtr hwnd = this.TestForm.Handle;

            using (reader) {
                while (reader.Read()) {
                    if (reader.NodeType == XmlNodeType.Whitespace || 
                        reader.NodeType == XmlNodeType.SignificantWhitespace ||
                        reader.NodeType == XmlNodeType.XmlDeclaration)
                        continue;

                    if (hwnd != Window.GetForegroundWindowHandle()) {
                        Sleep(1000);
                        throw new ApplicationException("Unexpected popup:" + Window.GetForegroundWindowText());
                    }

                    Trace.WriteLine(string.Format("Adding node type {0} with name {1} and value {2}",
                        reader.NodeType.ToString(), reader.Name, reader.Value));
                    
                    nodes.Add(new NodeInfo(reader));
                    bool children = false;
                    Trace.WriteLine(reader.NodeType + " " + reader.Name + "["+reader.Value+"]");
                    switch (reader.NodeType) {
                        case XmlNodeType.Element:
                            commands++;
                            TestForm.InvokeMenuItem(openElement ? "elementChildToolStripMenuItem" :
                                "elementAfterToolStripMenuItem");
                            openElement = true;
                            bool isEmpty = reader.IsEmptyElement;
                            if (!isEmpty) {
                                hasChildren.Push(children);
                                children = false;
                            } else {
                                openElement = false;
                            }
                            string name = reader.Name;
                            SendKeys.SendWait(name);
                            Sleep(20);
                            SendKeys.SendWait("{ENTER}");

                            readyForText = true;
                            bool firstAttribute = true;
                            while (reader.MoveToNextAttribute()){
                                TestForm.InvokeMenuItem(firstAttribute ? "attributeChildToolStripMenuItem" :
                                    "attributeAfterToolStripMenuItem");
                                firstAttribute = false;
                                readyForText = false;
                                openElement = false;
                                children = true;
                                SendKeys.SendWait(reader.Name + "{TAB}");
                                commands++;
                                SendKeys.SendWait(reader.Value);
                                Sleep(20);
                                SendKeys.SendWait("{ENTER}");
                                commands++;
                                Sleep(30);
                                SendKeys.SendWait("{LEFT}");
                            }
                            if (isEmpty) {
                                readyForText = false;
                                Sleep(50);
                                if (firstAttribute) SendKeys.SendWait("{ESC}"); // cancel value editing.
                                Sleep(30);
                                SendKeys.SendWait("{LEFT}");
                                Sleep(50);
                            }
                            break;
                        case XmlNodeType.Comment:
                            children = true;
                            TestForm.InvokeMenuItem(openElement ? "commentChildToolStripMenuItem" : "commentAfterToolStripMenuItem");
                            commands++;
                            SendKeys.SendWait(reader.Value);
                            Sleep(20);
                            SendKeys.SendWait("{ENTER}");
                            commands++;
                            Sleep(30);
                            SendKeys.SendWait("{LEFT}");
                            openElement = false;
                            break;
                        case XmlNodeType.CDATA:
                            children = true;
                            TestForm.InvokeMenuItem(openElement ? "cdataChildToolStripMenuItem" : "cdataAfterToolStripMenuItem");
                            commands++;
                            SendKeys.SendWait(reader.Value);
                            Sleep(20);
                            SendKeys.SendWait("{ENTER}");
                            commands++;
                            Sleep(30);
                            SendKeys.SendWait("{LEFT}");
                            openElement = false;
                            break;
                        case XmlNodeType.Text:
                            Sleep(30);
                            if (openElement) {
                                commands++;
                                if (!readyForText)
                                    SendKeys.SendWait("{TAB}{ENTER}");
                            } else {
                                children = true;
                                commands++;
                                TestForm.InvokeMenuItem("textAfterToolStripMenuItem");
                                openElement = false;
                            }
                            SendKeys.SendWait(reader.Value);
                            Sleep(20);
                            SendKeys.SendWait("{ENTER}");
                            readyForText = false;
                            commands++;
                            Sleep(30);
                            SendKeys.SendWait("{LEFT}");
                            break;
                        case XmlNodeType.ProcessingInstruction:
                            children = true;
                            TestForm.InvokeMenuItem(openElement ? "PIChildToolStripMenuItem" : "PIAfterToolStripMenuItem");
                            commands++;
                            SendKeys.SendWait(reader.Name + "{TAB}");
                            SendKeys.SendWait(reader.Value);
                            Sleep(20);
                            SendKeys.SendWait("{ENTER}");
                            commands++;
                            Sleep(30);
                            SendKeys.SendWait("{LEFT}");
                            openElement = false;
                            break;
                        case XmlNodeType.EndElement:
                            Sleep(50);
                            string keys = "";
                            if (readyForText) {
                                readyForText = false;
                                keys += "{ESC}{LEFT}";
                            }
                            if (children) {
                                keys += "{LEFT}";
                            }
                            children = hasChildren.Pop();
                            if (!openElement) {
                                keys += "{LEFT}";
                            }
                            SendKeys.SendWait(keys);
                            openElement = false;
                            Sleep(50);
                            break;
                    }
                }
            }

            // Test undo-redo
            UndoRedo(commands);

            this.SaveAndCompare("out.xml", "test1.xml");        
    
        }

        [TestMethod]
        public void TestEditCombinations() {
            Trace.WriteLine("TestEditCombinations==========================================================");
            // Test all the combinations of insert before, after, child stuff!
            string testFile = TestDir + "UnitTests\\test1.xml";
            this.LaunchNotepad(testFile);
            TestForm.InvokeMenuItem("newToolStripMenuItem");
            Sleep(500);

            // each node type at root level
            string[] nodeTypes = new string[]{ "comment", "PI", "element", "attribute", "text", "cdata" };
            bool[] validInRoot = new bool[]{ true, true, true, false, false, false };
            bool[] requiresName = new bool[]{ false, true, true, true, false, false };
            string[] clips = new string[] { "<!--{1}-->", "<?{0} {1}?>", "<{0}>{1}</{0}>", "{0}=\"{1}\"", "{1}", "<![CDATA[{1}]]>" };
            nodeIndex = 0;

            for (int i = 0; i < nodeTypes.Length; i++){
                string type = nodeTypes[i];
                if (validInRoot[i]){
                    InsertNode(type, "Child", requiresName[i], clips[i]);
                    Undo();
                    Undo();
                }
            }

            TestForm.InvokeMenuItem("commentChildToolStripMenuItem");

            for (int i = 0; i < nodeTypes.Length; i++) {
                string type = nodeTypes[i];
                if (validInRoot[i]) {
                    InsertNode(type, "After", requiresName[i], clips[i]);                                       
                    if (type != "element") {
                        InsertNode(type, "Before", requiresName[i], clips[i]);                    
                    }
                }
            }
            SendKeys.SendWait("^Ielement");
                
            // test all combinations of child elements under root element
            for (int i = 0; i < nodeTypes.Length; i++) {
                string type = nodeTypes[i];
                InsertNode(type, "Child", requiresName[i], clips[i]);
                InsertNode(type, "After", requiresName[i], clips[i]);
                InsertNode(type, "Before", requiresName[i], clips[i]);
                SendKeys.SendWait("{LEFT}{LEFT}"); // go back up to element.
            }
            this.SaveAndCompare("out.xml", "test7.xml");

        }

        int nodeIndex = 0;
        private void InsertNode(string type, string mode, bool requiresName, string clip) {
            string command = type + mode + "ToolStripMenuItem";
            Trace.WriteLine(command);
            TestForm.InvokeMenuItem(command);
            string name = type + nodeIndex.ToString();
            if (requiresName) {
                SendKeys.SendWait(name);
                SendKeys.SendWait("{TAB}");
            }
            string value = mode;
            SendKeys.SendWait(value + "{ENTER}");
            clip = string.Format(clip, name, value);
            TestForm.InvokeMenuItem("toolStripButtonCopy");
            CheckClipboard(clip);
            Clipboard.SetText("error");
            UndoRedo(2);
            TestForm.InvokeMenuItem("toolStripButtonCopy");
            CheckClipboard(clip);

            nodeIndex++;            
        }

        void CheckNodeValue(string expected) {
            Sleep(100);
            SendKeys.SendWait("{ENTER}");
            this.TestForm.WaitForIdle(1000);
            Sleep(100);
            SendKeys.SendWait("^c");
            CheckClipboard(expected);
            this.TestForm.WaitForIdle(1000);
            Sleep(300);
            SendKeys.SendWait("{ENTER}"); 
            Sleep(1000);
            //Sleep(1000);
        }

        internal void CheckOuterXml(string expected) {
            SendKeys.SendWait("^c");
            CheckClipboard(expected);
        }

        [TestMethod]
        public void TestIntellisense() {
            Trace.WriteLine("TestIntellisense==========================================================");
            LaunchNotepad();

            string outFile = TestDir + "UnitTests\\out.xml";

            Trace.WriteLine("Add <Basket>");
            TestForm.InvokeMenuItem("elementChildToolStripMenuItem");
            SendKeys.SendWait("Basket{ENTER}");

            Save("out.xml");

            Trace.WriteLine("Add xmlns:xsi attribute");
            TestForm.InvokeMenuItem("attributeChildToolStripMenuItem");
            SendKeys.SendWait("xmlns:xsi{TAB}");
            SendKeys.SendWait("http://www.w3.org/2001/XMLSchema-instance");

            Trace.WriteLine("Add xsi:noNamespaceSchemaLocation attribute");
            TestForm.InvokeMenuItem("attributeAfterToolStripMenuItem");
            SendKeys.SendWait("xsi:no{TAB}");

            Trace.WriteLine("code coverage on Checker ReportError for non-existant schema");
            SendKeys.SendWait("foo.xsd{ENTER}");
            Sleep(500);
            Trace.WriteLine("Point to xsd2.xsd");
            SendKeys.SendWait("{ENTER}test2.xsd{ENTER}");

            Trace.WriteLine("Get intellisense tooltip");
            SystemAccessible xtv = this.XmlTreeView;
            Sleep(2000); // wait for tooltip!
            Rectangle treeBounds = xtv.Bounds;
            Mouse.MouseMoveTo(Control.MousePosition, new Point(treeBounds.Left + 20, treeBounds.Top + 10), 5);
            Sleep(2000); // wait for tooltip!

            Trace.WriteLine("Add language='en-au'");
            TestForm.InvokeMenuItem("attributeAfterToolStripMenuItem");
            SendKeys.SendWait("l{TAB}en-a{ENTER}");

            Trace.WriteLine("Add sold='true'");
            TestForm.InvokeMenuItem("attributeAfterToolStripMenuItem");
            SendKeys.SendWait("s{TAB}t{ENTER}");

            Trace.WriteLine("Test validation error!");
            TestForm.InvokeMenuItem("attributeAfterToolStripMenuItem");
            SendKeys.SendWait("tick{TAB}12345{ENTER}");
            Sleep(500); //just so I can see it

            Trace.WriteLine("Test we can rename attributes!");
            SendKeys.SendWait("+{TAB}{ENTER}");
            SendKeys.SendWait("t{ENTER}");

            Trace.WriteLine("make sure we can undo redo edit attribute name."); 
            UndoRedo(); 
            SendKeys.SendWait("{TAB}{ENTER}11:18:38:P{ENTER}");

            Trace.WriteLine("undo redo of edit attribute");
            UndoRedo();

            Trace.WriteLine("Add <color>Thistle</color>");
            TestForm.InvokeMenuItem("elementAfterToolStripMenuItem");
            SendKeys.SendWait("c{TAB}ste{ENTER}");

            Trace.WriteLine("Check intellisense dropdown navigation.");
            CheckNodeValue("SteelBlue");

            Trace.WriteLine("Change Steeleblue to Yellow");
            SendKeys.SendWait("{ENTER}");

            SendKeys.SendWait("{PGDN}{END}{UP}{ENTER}");
            CheckNodeValue("Yellow");
            SendKeys.SendWait("{ENTER}{PGUP}{HOME}{DOWN}{DOWN}{ENTER}");
            CheckNodeValue("Aqua");

            Trace.WriteLine("Click Color Picker button.");
            Sleep(500);
            SendKeys.SendWait("{ENTER}"); // activate drop down.
            Rectangle bounds = ClickXmlBuilder();
            SendKeys.SendWait("{DOWN}{LEFT} {ENTER}");

            Trace.WriteLine("test MouseDown on NodeTextView editor");
            Mouse.MouseClick(new Point(bounds.Left + 20, bounds.Top - 10), MouseButtons.Left);
            Sleep(500);
            SendKeys.SendWait("{DOWN}{ENTER}");

            Trace.WriteLine("Add <date>2005-12-25</date>");
            TestForm.InvokeMenuItem("elementAfterToolStripMenuItem");
            SendKeys.SendWait("d{TAB}12/25/2005{ENTER}");

            Trace.WriteLine("Add <dateTime>2006-11-11T07:30:00</dateTime> ");
            TestForm.InvokeMenuItem("elementAfterToolStripMenuItem");
            SendKeys.SendWait("d{TAB}11/11/2006{RIGHT}07:30:00:A{ENTER}");

            Trace.WriteLine("Add <photo>test1.xml</photo> ");
            TestForm.InvokeMenuItem("elementAfterToolStripMenuItem");
            SendKeys.SendWait("p{TAB}basket.jpg");

            Trace.WriteLine("Test UriBuilder");
            ClickXmlBuilder();
            SendKeys.SendWait(TestDir + "UnitTests\\" + "test1.xml");
            SendKeys.SendWait("{ENTER}");

            Trace.WriteLine("Add <fruit>banana</fruit> ");
            TestForm.InvokeMenuItem("elementAfterToolStripMenuItem");
            SendKeys.SendWait("fr{TAB}b{ENTER}");

            Trace.WriteLine("Add <font>Arial, 8.25</font> ");
            TestForm.InvokeMenuItem("elementAfterToolStripMenuItem");
            SendKeys.SendWait("fo{TAB}ar{ENTER}");
            Sleep(500);//just so I can see it

            Trace.WriteLine("Add Test FontBuilder");
            SendKeys.SendWait("{ENTER}");
            ClickXmlBuilder();
            SendKeys.SendWait("{ENTER}{ENTER}");

            Trace.WriteLine("Add <vegetable>cucumber</vegetable> ");
            TestForm.InvokeMenuItem("elementAfterToolStripMenuItem");
            SendKeys.SendWait("v{TAB}cu{ENTER}");

            Trace.WriteLine("Add <berry>huckleberry</berry> ");
            TestForm.InvokeMenuItem("elementAfterToolStripMenuItem");
            SendKeys.SendWait("b{TAB}hu{ENTER}");

            Trace.WriteLine("Test edit of PI name");
            TestForm.InvokeMenuItem("PIAfterToolStripMenuItem");
            SendKeys.SendWait("test{ENTER}{ESC}{LEFT}");
            Sleep(500);//just so I can see it
            SendKeys.SendWait("{ENTER}pi{ENTER}");
            Sleep(500);//just so I can see it
            UndoRedo();

            Trace.WriteLine("Test validation error and elementBefore command!");
            TestForm.InvokeMenuItem("elementBeforeToolStripMenuItem");
            SendKeys.SendWait("woops{ENTER}");
            Sleep(500);//just so I can see it

            Trace.WriteLine("Move to Basket element"); 
            SendKeys.SendWait("{ESC}");            
            Sleep(500);
            SendKeys.SendWait("{LEFT}{LEFT}");
            Sleep(1000);
            Trace.WriteLine("Navigate to next error");
            NavigateNextError();
            CheckNodeValue("woops");
            Trace.WriteLine("Move to Basket element"); 
            SendKeys.SendWait("{LEFT}"); 

            Trace.WriteLine("Navigate error with mouse double click");
            NavigateErrorWithMouse();

            Trace.WriteLine("We are now back on the 'woops' element.");
            CheckNodeValue("woops");            

            Trace.WriteLine("undo redo of elementBeforeToolStripMenuItem.");
            UndoRedo();

            Trace.WriteLine("Add <weight>1.234</weight> ");
            SendKeys.SendWait("{TAB}{ENTER}1.234{ENTER}");

            Trace.WriteLine("Test we can fix it by renaming element");
            SendKeys.SendWait("+{TAB}{ENTER}w{TAB}");

            Trace.WriteLine("undo redo of edit element name");
            UndoRedo();

            this.SaveAndCompare("out.xml", "test2.xml");
        }

        private void NavigateErrorWithMouse() {
            AccessibleObject grid = TestForm.FindDescendant("DataGridView");
            AccessibleObject row = grid.Navigate(AccessibleNavigation.FirstChild);
            row = row.Navigate(AccessibleNavigation.Next);
            Point pt = Center(row.Bounds);
            // Double click it
            Mouse.MouseDoubleClick(pt, MouseButtons.Left);
        }

        private void NavigateNextError() {
            TestForm.InvokeMenuItem("nextErrorToolStripMenuItem");
        }

        private void Undo(int count) {
            while (count-- > 0) {
                Undo();
            }
        }

        private void Redo(int count) {
            while (count-- > 0) {
                Redo();
            }
        }

        private void Undo() {
            TestForm.InvokeMenuItem("undoToolStripMenuItem");
        }
        private void Redo() {
            TestForm.InvokeMenuItem("redoToolStripMenuItem");
        }
        private void UndoRedo(int level) {
            for (int i = 0; i < level; i++) {
                Undo();
            }
            for (int i = 0; i < level; i++) {
                Redo();
            }
        }
        private void UndoRedo() {
            Undo();
            Redo();
        }

        Rectangle ClickXmlBuilder() {
            // Find the intellisense button and click on it
            Rectangle bounds = NodeTextViewCompletionSet.Bounds;
            Sleep(1000);
            Mouse.MouseClick(new Point(bounds.Left + 15, bounds.Top + 10), MouseButtons.Left);
            TestForm.WaitForNewWindow();
            return bounds;
        }


        [TestMethod]
        public void TestCompare() {
            Trace.WriteLine("TestCompare==========================================================");
            string testFile = TestDir + "UnitTests\\test4.xml";
            LaunchNotepad(testFile);

            // something the same
            TestForm.InvokeAsyncMenuItem("compareXMLFilesToolStripMenuItem");
            TestForm.WaitForPopup();
            SendKeys.SendWait(TestDir + "UnitTests\\test4.xml{ENTER}");
            TestForm.WaitForNewWindow();
            string text = Window.GetForegroundWindowText();
            Assert.AreEqual<string>(text, "Files Identical");
            SendKeys.SendWait("{ENTER}");

            // Now something different
            TestForm.InvokeAsyncMenuItem("compareXMLFilesToolStripMenuItem");
            
            TestForm.WaitForPopup();            
            SendKeys.SendWait(TestDir + "UnitTests\\test5.xml{ENTER}");
            TestForm.WaitForNewWindow();
            SendKeys.SendWait("%{F4}");

            Undo();
        }        

        [TestMethod]
        public void TestClipboard() {
            Trace.WriteLine("TestClipboard==========================================================");

            string testFile = TestDir + "UnitTests\\test1.xml";
            LaunchNotepad(testFile);

            Trace.WriteLine("Incremental find 'Emp'");
            TestForm.InvokeMenuItem("incrementalSearchToolStripMenuItem");
            SendKeys.SendWait("Emp");

            Trace.WriteLine("cut");
            SendKeys.SendWait("^x");
            
            string expected = "<Employee xmlns=\"http://www.hr.org\" id=\"46613\" title=\"Architect\"><Name First=\"Chris\" Last=\"Lovett\" /><Street>One Microsoft Way</Street><City>Redmond</City><Zip>98052</Zip><Country><Name>U.S.A.</Name></Country><Office /></Employee>";
            CheckClipboard(expected);

            Trace.WriteLine("paste something different");
            string expected2 = "<Employee xmlns=\"http://www.hr.org\" id=\"46613\" title=\"Architect\"><Name>Test</Name><?pi test?></Employee>";
            SendKeys.SendWait("{LEFT}");
            Clipboard.SetText(expected2);
            SendKeys.SendWait("^v");
            Sleep(500);
            CheckOuterXml(expected2);

            Trace.WriteLine("test undo and redo of cut/paste commands");
            UndoRedo(2);
            Undo();
            Undo();

            CheckOuterXml(expected);

            Trace.WriteLine("test delete key");
            TestForm.InvokeMenuItem("deleteToolStripMenuItem");
            CheckOuterXml("<!--last comment-->");
            UndoRedo();
            Undo();

            Trace.WriteLine("test copy paste via menus");
            SendKeys.SendWait("{END}"); // now on #comment
            Clipboard.SetText("error");
            TestForm.InvokeMenuItem("copyToolStripMenuItem");
            CheckClipboard("<!--last comment-->");
            Clipboard.SetText("error");
            TestForm.InvokeMenuItem("cutToolStripMenuItem");
            CheckClipboard("<!--last comment-->");
            TestForm.InvokeMenuItem("pasteToolStripMenuItem");
            Undo();
            Undo();

            Trace.WriteLine("Test repeat");
            TestForm.InvokeMenuItem("repeatToolStripMenuItem");
            SendKeys.SendWait("new comment{ENTER}");
            UndoRedo(2); // test redo of duplicate!
            Undo();
            Undo();

            Trace.WriteLine("test cut/copy/paste/delete in NodeTextView");
            SendKeys.SendWait("{DEL}");
            SendKeys.SendWait("^z");
            CheckNodeValue("last comment");
            CheckOuterXml("<!--last comment-->");
            SendKeys.SendWait("^x");
            CheckClipboard("<!--last comment-->");
            Undo();
            Clipboard.SetText("<!--last comment-->");
            SendKeys.SendWait("^v");
            CheckOuterXml("<!--last comment-->");
            Undo();
            Sleep(5000);
            Trace.WriteLine("type to find in 'foo' in node text view");
            SendKeys.SendWait("^Ifo");
            CheckNodeValue("foo");

            Trace.WriteLine("DuplicateNode");
            SendKeys.SendWait("{TAB}^IEmp");
            CheckNodeValue("Employee");
            TestForm.InvokeMenuItem("duplicateToolStripMenuItem");

            Trace.WriteLine("undo/redo of duplicate");
            UndoRedo();
            Undo();
            CheckNodeValue("Employee");

            Trace.WriteLine("Test namespace aware copy/paste");
            string xml = "<x:item xmlns:x='uri:1'>Some text</x:item>";
            Clipboard.SetText(xml);
            TestForm.InvokeMenuItem("pasteToolStripMenuItem");
            Sleep(500);

            Trace.WriteLine("Test namespace normalization on paste.");
            TestForm.InvokeMenuItem("pasteToolStripMenuItem");

            Trace.WriteLine("Test namespace prefix auto-generation.");
            Sleep(500); 
            SendKeys.SendWait("{DOWN}"); // reset type-to-find
            SendKeys.SendWait("^Iid");
            SendKeys.SendWait("{ENTER}");
            SendKeys.SendWait("{HOME}y:{ENTER}");
            Sleep(200); 
            SendKeys.SendWait("{DOWN}"); // reset type-to-find
            SendKeys.SendWait("^Iitem");
            SendKeys.SendWait("{ENTER}{HOME}z:{ENTER}");
            Sleep(200);
            
            // test save to same file.
            this.SaveAndCompare("out.xml", "test6.xml");

        }

        void WipeFile(string fname) {
            if (File.Exists(fname)) {
                File.SetAttributes(fname, File.GetAttributes(fname) & ~FileAttributes.ReadOnly);
                File.Delete(fname);
            }
        }


        [TestMethod]
        public void TestOptionsDialog() {
            Trace.WriteLine("TestOptionsDialog==========================================================");
            
            // Save original settings.
            string originalSettings = Environment.GetEnvironmentVariable("USERPROFILE") + "\\Local Settings\\Application Data\\Microsoft\\Xml Notepad\\XmlNotepad.settings";
            string backupSettings = Path.GetTempPath() + "XmlNotepad.settings";
            File.Copy(originalSettings, backupSettings, true);

            LaunchNotepad();

            // Options dialog
            Trace.WriteLine("Options dialog...");
            Window options = TestForm.OpenDialog("optionsToolStripMenuItem", "FormOptions");

            // Find the PropertyGrid control.
            SystemAccessible acc = options.FindDescendant("propertyGrid1");

            SystemAccessible properties = acc.FindChild("Properties Window");
            SystemAccessible table = properties.FindChild("Properties Window");

            Trace.WriteLine("Font");
            SystemAccessible font = table.FindChild("Font"); // this is the group heading
            font = font.NextSibling;  // now get the actual editable font row.
            font.Select(AccessibleSelection.TakeSelection);
            Rectangle r = font.Bounds;
            // bring up the font dialog.
            Mouse.MouseClick(new Point(r.Right - 10, r.Top + 6), MouseButtons.Left);
            options.WaitForPopup();

            options.DismissPopUp("{ENTER}");

            string[] names = new string[] { "Element", "Attribute", "Text",
                    "Background", "Comment", "PI", "CDATA" };

            string[] values = new string[] { "Aqua", "128, 64, 64", "64, 0, 0",
                  "64, 0, 128", "Lime", "128, 0, 64", "0, 64, 64"};

            
            for (int i = 0, n = names.Length; i<n ;i++) {
                string name = names[i];

                Trace.WriteLine("Click " + name);
                
                AccessibleObject child = table.FindChild(name);
                child.Select(AccessibleSelection.TakeSelection);                
                SendKeys.SendWait("{TAB}" + values[i] + "{ENTER}");

                Sleep(333); // so we can see it!
            }

            this.TestForm.DismissPopUp("%O");            
            bool passed = true;

            // Close the app.
            TestForm.Dispose();
            Sleep(1000); // give it time to write out the new setrtings.

            // verify persisted colors.
            if (File.Exists(originalSettings)) {
                XmlDocument doc = new XmlDocument();
                doc.Load(originalSettings);
                for (int i = 0, n = names.Length; i<n; i++){
                    string ename = names[i];
                    XmlNode node = doc.SelectSingleNode("Settings/Colors/" + ename);
                    if (node != null) {
                        string expected = values[i];
                        string actual = node.InnerText;
                        if (expected != actual) {
                            Trace.WriteLine(string.Format("Color '{0}' has unexpected value '{1}'", ename, actual));
                            passed = false;
                        }
                    }
                }
            }

            // restore the original settings.
            File.Copy(backupSettings, originalSettings, true);
            DeleteFile(backupSettings);

            if (!passed) {
                throw new ApplicationException("Unexpected colors found in XmlNotepad.settings file.");
            }
            
        }

        [TestMethod]
        public void TestDialogs() {
            Trace.WriteLine("TestDialogs==========================================================");
            LaunchNotepad();

            // About...
            Trace.WriteLine("About...");
            TestForm.InvokeAsyncMenuItem("aboutXMLNotepadToolStripMenuItem");
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("{ENTER}");

            // hide/show status bar
            Trace.WriteLine("hide/show status bar...");
            TestForm.InvokeAsyncMenuItem("statusBarToolStripMenuItem");
            Sleep(500);
            TestForm.InvokeAsyncMenuItem("statusBarToolStripMenuItem");
            Sleep(500);

            // open bad file.            
            Trace.WriteLine("open bad file");
            TestForm.InvokeAsyncMenuItem("openToolStripMenuItem");
            TestForm.WaitForPopup();
            SendKeys.SendWait(TestDir + "UnitTests\\bad.xml{ENTER}");
            TestForm.WaitForPopup();
            SendKeys.SendWait("%Y");
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("%{F4}");

            // Test OpenFileDialog
            Trace.WriteLine("OpenFileDialog");
            TestForm.InvokeAsyncMenuItem("openToolStripMenuItem");
            TestForm.WaitForPopup();
            SendKeys.SendWait(TestDir + "UnitTests\\supply.xml");
            TestForm.DismissPopUp("{ENTER}");

            Trace.WriteLine("Test long line wrap message.");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            TestForm.WaitForPopup();
            SendKeys.SendWait("FinalDeliverable{ENTER}");            
            TestForm.DismissPopUp("{ESC}");

            Trace.WriteLine("Oh, and test horizontal scroll bar");
            SystemAccessible hscroll = XmlTreeView.FindChild("HScrollBar");
            Rectangle sbBounds = hscroll.Bounds;
            Sleep(1000);
            Mouse.MouseClick(new Point(sbBounds.Left + 5, sbBounds.Top + 5), MouseButtons.Left);
            Sleep(500);

            SendKeys.SendWait("{TAB}{ENTER}");
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("{ENTER}");
            SendKeys.SendWait("{ENTER}");
                        
            Trace.WriteLine("View source");
            TestForm.InvokeAsyncMenuItem("sourceToolStripMenuItem");
            TestForm.WaitForPopup(); // file has changed, do you want to save it?
            SendKeys.SendWait("%N");
            TestForm.WaitForPopup(); // wait for Notepad.exe.
            TestForm.DismissPopUp("%{F4}");
            
            // Show help
            Trace.WriteLine("Show help...");
            Trace.WriteLine(Directory.GetCurrentDirectory());
            Trace.WriteLine(Application.StartupPath);
            SendKeys.SendWait("{F1}");
            TestForm.WaitForNewWindow();
            TestForm.DismissPopUp("%{F4}");
            
            // Test reload - discard changes
            Trace.WriteLine("Reload- discard changes");
            TestForm.InvokeAsyncMenuItem("reloadToolStripMenuItem");
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("{ENTER}");

            // Save As...
            Trace.WriteLine("Save As..."); 
            string outFile = TestDir + "UnitTests\\out.xml";
            WipeFile(outFile);
            TestForm.InvokeAsyncMenuItem("saveAsToolStripMenuItem");
            TestForm.WaitForPopup();
            SendKeys.SendWait("out.xml");
            TestForm.DismissPopUp("{ENTER}");

            // Check save read only
            Trace.WriteLine("Check save read only.");
            File.SetAttributes(outFile, File.GetAttributes(outFile) | FileAttributes.ReadOnly);
            TestForm.InvokeAsyncMenuItem("saveToolStripMenuItem");
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("%Y");
            Sleep(2000); // let file system settle down...
            
            // Test "reload" message box.
            Trace.WriteLine("File has changed on disk, do you want to reload?");
            File.SetLastWriteTime(outFile, DateTime.Now);
            Sleep(2000); // now takes 2 seconds for this to show up.
            
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("%Y"); // reload!
            
            // Window/NewWindow!
            Trace.WriteLine("Window/NewWindow");
            TestForm.InvokeAsyncMenuItem("newWindowToolStripMenuItem");
            TestForm.WaitForNewWindow();

            TestForm.DismissPopUp("%{F4}"); // close second window!
            
            if (!Window.GetForegroundWindowText().StartsWith("XML Notepad")) {
                TestForm.Activate(); // alt-f4 sometimes sends focus to another window (namely, the VS process running this test!)
                Sleep(500);
            }
            Sleep(1000);
           
            // Test SaveIfDirty
            Trace.WriteLine("make simple edit");
            FocusTreeView();
            SendKeys.SendWait("{END}{DELETE}");// make simple edit
           
            Trace.WriteLine("Test error dialog when user tries to enter element with no name");
            TestForm.InvokeMenuItem("repeatToolStripMenuItem");
            SendKeys.SendWait("{ENTER}");
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("%N");
            
            Trace.WriteLine("Test error dialog when user tries to enter name with spaces");
            SendKeys.SendWait("     {ENTER}");
            TestForm.WaitForPopup();
            Trace.WriteLine("This time accept the empty name");
            TestForm.DismissPopUp("%Y");

            Trace.WriteLine("Test error dialog when user enter an invalid name");
            SendKeys.SendWait("{ENTER}{+}{+}{+}{ENTER}");
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("{ENTER}");
            SendKeys.SendWait("{ESC}");
            Undo();
            
            // Save changes on exit?
            Trace.WriteLine("Save changes on exit - cancel");
            TestForm.InvokeAsyncMenuItem("exitToolStripMenuItem");
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("{ESC}"); // make sure we can cancel exit!
            Sleep(1000);

            // Save changes on 'new'?
            Trace.WriteLine("Save changes on 'new' - cancel");
            TestForm.InvokeAsyncMenuItem("newToolStripMenuItem");
            TestForm.WaitForPopup();
            TestForm.DismissPopUp("{ESC}"); // make sure we can cancel 'new'!
            Sleep(1000);

            Trace.WriteLine("Save changes on 'exit' - yes!");
            CheckNodeName("Header");
            TestForm.InvokeAsyncMenuItem("exitToolStripMenuItem");
            TestForm.WaitForPopup();
            SendKeys.SendWait("%Y"); // save the changes!

        }

        [TestMethod]
        public void TestSchemaDialog() {
            Trace.WriteLine("TestSchemaDialog==========================================================");
            LaunchNotepad();

            Sleep(1000);
            Trace.WriteLine("Open Schema Dialog");
            Window schemaDialog = TestForm.OpenDialog("schemasToolStripMenuItem", "FormSchemas");
            schemaDialog.InvokeMenuItem("clearToolStripMenuItem");

            Trace.WriteLine("Add schema via file dialog");
            SendKeys.SendWait("{TAB}{TAB}{TAB} "); // bring up file open dialog
            Sleep(1000);
            schemaDialog.WaitForPopup();
            string schema = TestDir + "UnitTests\\emp.xsd";
            SendKeys.SendWait(schema);
            schemaDialog.DismissPopUp("{ENTER}");
            
            SendKeys.SendWait("^{HOME}+ "); // select first row
            Sleep(300); // just so we can watch it happen
            SendKeys.SendWait("^c"); // copy
            string text = Clipboard.GetText();
            if (!text.ToLowerInvariant().Contains("emp.xsd")) {
                throw new ApplicationException("Did not find 'test2.xsd' on the clipboard!");
            }
            Trace.WriteLine("Close schema dialog");
            TestForm.DismissPopUp("%O"); // hot key for OK button.
 
            Trace.WriteLine("Close XML Notepad and reload it to ensure schema cache was persisted");
            CloseApp();
            LaunchNotepad();
            schemaDialog = TestForm.OpenDialog("schemasToolStripMenuItem", "FormSchemas");
            Sleep(500);
            SendKeys.SendWait("^{HOME}+ "); // select first row

            Trace.WriteLine("Cut");
            SendKeys.SendWait("^x"); // cut
            text = Clipboard.GetText();
            if (!text.ToLowerInvariant().Contains("emp.xsd")) {
                throw new ApplicationException("Did not find 'test2.xsd' on the clipboard!");
            }
            Sleep(300);
            Trace.WriteLine("Paste");
            SendKeys.SendWait("^v"); // paste
            Sleep(300);

            Trace.WriteLine("Edit of filename cell.");
            SendKeys.SendWait("^{HOME}{RIGHT}{RIGHT}" + schema + "{ENTER}");
            Trace.WriteLine("Undo");
            SendKeys.SendWait("^z"); // undo
            Sleep(300);
            Trace.WriteLine("Redo");
            SendKeys.SendWait("^y"); // redo            
            Sleep(300);
            Trace.WriteLine("Delete");
            SendKeys.SendWait("^{HOME}+ {DELETE}"); // delete first row
            Sleep(300);
            Trace.WriteLine("Undo");
            SendKeys.SendWait("^z"); // undo
            Sleep(300);
            Trace.WriteLine("Redo");
            SendKeys.SendWait("^y"); // redo
            Sleep(300);
            Trace.WriteLine("Undo");
            SendKeys.SendWait("^z"); // undo
            Sleep(300);

            Trace.WriteLine("Make sure we commit with some rows to update schema cache!");
            TestForm.DismissPopUp("%O"); // hot key for OK button.

            Trace.WriteLine("Now add a duplicate target namespcace.");
            schemaDialog = TestForm.OpenDialog("schemasToolStripMenuItem", "FormSchemas");
            Sleep(500);

            Trace.WriteLine("Add emp2.xsd via paste");
            schema = TestDir + "UnitTests\\emp2.xsd";
            SendKeys.SendWait("{DOWN}{RIGHT}{RIGHT}^ "); // select first row
            Clipboard.SetText(schema);
            SendKeys.SendWait("^v");
            SendKeys.SendWait("^c"); // copy
            text = Clipboard.GetText();
            if (!text.ToLowerInvariant().Contains("emp2.xsd")) {
                throw new ApplicationException("Did not find 'test2.xsd' on the clipboard!");
            }

            Trace.WriteLine("Add duplicate schema via file dialog ");
            Sleep(1000);
            schemaDialog.InvokeAsyncMenuItem("addSchemasToolStripMenuItem");
            TestForm.WaitForPopup();
            schema = TestDir + "UnitTests\\emp.xsd";
            SendKeys.SendWait(schema + "{ENTER}");
            
            Sleep(300); // just so we can watch it happen
            SendKeys.SendWait("^c"); // copy
            text = Clipboard.GetText();
            if (!text.ToLowerInvariant().Contains("emp.xsd")) {
                throw new ApplicationException("Did not find 'emp.xsd' on the clipboard!");
            }
            Trace.WriteLine("Make sure we commit with the duplicate tns.");
            SendKeys.SendWait("%O"); // hot key for OK button.

            Trace.WriteLine("Clear the schemas.");
            schemaDialog = TestForm.OpenDialog("schemasToolStripMenuItem", "FormSchemas");
            schemaDialog.InvokeMenuItem("clearToolStripMenuItem");
            // Make sure we commit with the duplicate tns.
            SendKeys.SendWait("%O"); // hot key for OK button.

        }

        string CopyXPathFindString(Window w) {
            SystemAccessible s = w.FindDescendant("checkBoxXPath");
            if (s.DefaultAction == "Check") {
                s.DoDefaultAction();
            }
            SystemAccessible c = w.FindDescendant("comboBoxFind");
            return c.Value;
        }

        [TestMethod]
        public void TestXPathFind() {
            Trace.WriteLine("TestXPathFind==========================================================");
            // Give view source something to show.
            string testFile = TestDir + "UnitTests\\test1.xml";
            LaunchNotepad(testFile);
                       
            Sleep(1000);

            Trace.WriteLine("test path of 'pi' node");
            TestForm.SendKeys("^Ipi");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            Window fd = TestForm.WaitForPopup();            
            AssertNormalizedEqual(CopyXPathFindString(fd), "/processing-instruction('pi')"); // test pi
            TestForm.DismissPopUp("{ESC}");

            Trace.WriteLine("test path of comment");
            TestForm.SendKeys("{DOWN}^I#");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            fd = TestForm.WaitForPopup();
            AssertNormalizedEqual(CopyXPathFindString(fd), "/Root/comment()[1]");
            TestForm.DismissPopUp("{ESC}");

            Trace.WriteLine("test path of cdata");
            TestForm.SendKeys("{ESC}{DOWN}");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            fd = TestForm.WaitForPopup();
            AssertNormalizedEqual(CopyXPathFindString(fd), "/Root/text()");
            TestForm.DismissPopUp("{ESC}");

            Trace.WriteLine("test path of text node");
            TestForm.SendKeys("{DOWN}{RIGHT}{DOWN}");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            fd = TestForm.WaitForPopup();
            AssertNormalizedEqual(CopyXPathFindString(fd), "/Root/item/text()");
            TestForm.DismissPopUp("{ESC}");

            Trace.WriteLine("test path of node with namespace");
            TestForm.SendKeys("{ESC}{DOWN}^IEmp");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            fd = TestForm.WaitForPopup();
            AssertNormalizedEqual(CopyXPathFindString(fd), "/Root/a:Employee"); // test element with namespaces!
            
            Trace.WriteLine("test edit path and find node.");
            fd.SendKeys("/Root{ENTER}"); 
            TestForm.DismissPopUp("{ESC}");

            Trace.WriteLine("test 'id' attribute path generation.");
            TestForm.SendKeys("{ESC}{DOWN}");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            fd = TestForm.WaitForPopup();
            AssertNormalizedEqual(CopyXPathFindString(fd), "/Root/@id");
            TestForm.DismissPopUp("{ESC}");

            Trace.WriteLine("Find on an xmlns attributue!");
            TestForm.SendKeys("^IEmp{RIGHT}{DOWN}");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            fd = TestForm.WaitForPopup();
            AssertNormalizedEqual(CopyXPathFindString(fd), "/Root/a:Employee/namespace::*[local-name()='']");
            TestForm.DismissPopUp("{ESC}");

            // XmlDocument lazily creates namespace nodes causing this test to "modify" the document!
            Save("out.xml");
        }

        [TestMethod]
        public void TestXsltOutput() {
            Trace.WriteLine("TestXsltOutput==========================================================");
                        
            LaunchNotepad();

            Trace.WriteLine("Click in the combo box location field");
            SystemAccessible comboBoxLocation = this.TestForm.FindDescendant("comboBoxLocation");
            Rectangle bounds = comboBoxLocation.Bounds;
            Mouse.MouseClick(Center(bounds), MouseButtons.Left);

            Trace.WriteLine("Load RSS from http");
            SendKeys.SendWait("{END}+{HOME}http://msdn.microsoft.com/rss.xml{ENTER}");

            Trace.WriteLine("Wait for rss to be loaded");
            WaitForText("<?xml version=\"1.0\" encoding=\"utf-8\"?>");

            SendKeys.SendWait("{DOWN}");
            this.CheckOuterXml("<?xml-stylesheet type='text/xsl' href='rsspretty.xsl' version='1.0'?>");

            Trace.WriteLine("Show XSLT");
            AccessibleObject tab = TestForm.FindDescendant("tabPageHtmlView");
            bounds = tab.Bounds;
            Trace.WriteLine("Select tabPageHtmlView ");
            Mouse.MouseClick(new Point(bounds.Left + 20 + 70, bounds.Top - 15), MouseButtons.Left);
            Sleep(1000);

            Trace.WriteLine("Enter custom XSL with script code.");
            this.EnterXslFilename(TestDir + "UnitTests\\rss.xsl");
            TestForm.WaitForPopup();
            string title = Window.GetForegroundWindowText();
            if (title != "Untrusted Script Code") {
                throw new ApplicationException("Expecting script security dialog");
            }
            Sleep(1000);
            TestForm.DismissPopUp("%Y");

            Trace.WriteLine("Make sure it executed");
            CopyHtml();
            this.CheckClipboard("Found 99 RSS items. The script executed successfully.");

            Trace.WriteLine("Try xslt with error");
            this.EnterXslFilename(TestDir + "UnitTests\\bad.xsl");
            Sleep(2000);            
            CopyHtml();
            this.CheckClipboard(@"Prefix 'user' is not defined. ");

            Trace.WriteLine("Back to tree view");
            tab = TestForm.FindDescendant("tabPageTreeView");
            Mouse.MouseClick(new Point(bounds.Left + 20, bounds.Top - 15), MouseButtons.Left);

            Sleep(1000);
            Save("out.xml");
        }

        void WaitForText(string value) {            
            int retries = 20;
            string clip = null;
            while (retries-- > 0) {
                SendKeys.SendWait("^c");
                clip = Clipboard.GetText();
                Trace.WriteLine("clip=" + clip);
                if (clip == value)
                    return;
                Sleep(2000);
            }
            throw new Exception("Not finding expected text '" + value + "', instead we got '" + clip + "'");
        }

        void EnterXslFilename(string filename) {
            SystemAccessible s = TestForm.FindDescendant("SourceFileName");
            Rectangle bounds = s.Bounds;
            Mouse.MouseClick(Center(bounds), MouseButtons.Left);
            Sleep(500);
            SendKeys.SendWait("{END}+{HOME}" + filename + "{ENTER}");            
        }

        string CopyHtml() {
            AccessibleObject xsltViewer = this.TestForm.FindDescendant("xsltViewer");
            Rectangle bounds = xsltViewer.Bounds;
            // Tripple click to select all the text
            Mouse.MouseClick(Center(bounds), MouseButtons.Left);
            Mouse.MouseClick(Center(bounds), MouseButtons.Left);
            Mouse.MouseClick(Center(bounds), MouseButtons.Left);
            Sleep(1000);
            SendKeys.SendWait("^c");
            return Clipboard.GetText();
        }

        void ClearFindCheckBoxes(Window findDialog) {
            foreach (string name in new string[] { "checkBoxMatchCase", "checkBoxWholeWord", "checkBoxRegex", "checkBoxXPath" }) {
                SystemAccessible checkBox = findDialog.FindDescendant(name);
                if ((checkBox.State & AccessibleStates.Checked) != 0) {
                    checkBox.DoDefaultAction();
                }            
            }
        }

        [TestMethod]
        public void TestFindReplace() {

            ResetFindOptions();

            Trace.WriteLine("TestFindReplace==========================================================");
            // Give view source something to show.
            string testFile = TestDir + "UnitTests\\test1.xml";
            LaunchNotepad(testFile);

            Trace.WriteLine("Test auto-move of Find Window to reveal what was found");
            Rectangle treeBounds = this.XmlTreeView.Bounds;
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            Window findDialog = TestForm.WaitForPopup();

            ClearFindCheckBoxes(findDialog);

            Rectangle findBounds = findDialog.GetScreenBounds();
            Point treeCenter = Center(treeBounds);
            Point findCenter = Center(findBounds);
            Point start = new Point(findBounds.Left + (findBounds.Width / 2), findBounds.Top + 15);
            Point end = new Point(start.X + treeCenter.X - findCenter.X, 
                                  start.Y + treeCenter.Y - findCenter.Y);
            Mouse.MouseClick(start, MouseButtons.Left);
            Mouse.MouseDragTo(start, end, 5, MouseButtons.Left);

            // Refocus the combo box...
            Sleep(500);
            SystemAccessible findCombo = findDialog.FindDescendant("comboBoxFind");
            Mouse.MouseClick(Center(findCombo.Bounds), MouseButtons.Left);
            
            Sleep(500);
            findDialog.SendKeys("Some{ENTER}");
            Sleep(500);
            
            TestForm.DismissPopUp("{ESC}");
            SendKeys.SendWait("^c{ESC}");
            CheckClipboard("Some");
            Sleep(200);
            SendKeys.SendWait("^{HOME}");
            
            Trace.WriteLine("Test find error dialog");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            findDialog = TestForm.WaitForPopup();
            SendKeys.SendWait("will not find{ENTER}");
            TestForm.ExpectingPopup("Find Error");
            findDialog.DismissPopUp("{ENTER}");

            Trace.WriteLine("test we can find the 'this' text.");
            SendKeys.SendWait("this{ENTER}");
            Sleep(200);
            TestForm.DismissPopUp("{ESC}");
            Sleep(200);
            SendKeys.SendWait("{ESC}");
            Sleep(200);
            SendKeys.SendWait("^c");
            CheckClipboard("<!-- This tests all element types -->");
            
            Trace.WriteLine("repeat find with shortcut");
            SendKeys.SendWait("{F3}{ESC}");
            Sleep(200);
            SendKeys.SendWait("^c");
            CheckClipboard(@"
    The XML markup in this version is Copyright © 1999 Jon Bosak.
    This work may freely be distributed on condition that it not be
    modified or altered in any way.
    ");

            Trace.WriteLine("Test illegal regular expressions.");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            findDialog = TestForm.WaitForPopup();
            SendKeys.SendWait("\\%e{ENTER}");
            findDialog.ExpectingPopup("Find Error");
            findDialog.DismissPopUp("{ENTER}");
            Sleep(500);

            Trace.WriteLine("Find 'Microsoft' using regular expressions.");
            SendKeys.SendWait("M[{^} ]*t");
            SendKeys.SendWait("{ENTER}{ESC}");
            Sleep(500);
            SendKeys.SendWait("{ESC}");
            Sleep(500);
            SendKeys.SendWait("^c");
            CheckClipboard("One Microsoft Way");

            Trace.WriteLine("test we can find 'last' in a comment only.");
            SendKeys.SendWait("{HOME}");
            TestForm.InvokeAsyncMenuItem("findToolStripMenuItem");
            findDialog = TestForm.WaitForPopup();
            SendKeys.SendWait("last%e{TAB}{TAB}c");
            SendKeys.SendWait("{ENTER}{ESC}");
            Sleep(500);
            SendKeys.SendWait("{ESC}");
            Sleep(500);
            SendKeys.SendWait("^c");
            CheckClipboard("<!--last comment-->");

            SendKeys.SendWait("{HOME}");
            TestForm.InvokeAsyncMenuItem("replaceToolStripMenuItem");
            findDialog = TestForm.WaitForPopup();

            Trace.WriteLine("Toggle dialog using ctrl+f & ctrl+h");
            SendKeys.SendWait("^f");            
            Sleep(300); // so I can see it...
            SendKeys.SendWait("^h");            
            Sleep(300);

            Trace.WriteLine("test we can replace 'This' using case sensitive.");
            SendKeys.SendWait("This{TAB}xxx%m%w{TAB}{TAB}{TAB}e%a");
            string expected = @"
    The XML markup in this version is Copyright © 1999 Jon Bosak.
    xxx work may freely be distributed on condition that it not be
    modified or altered in any way.
    ";
            Sleep(200);
            SendKeys.SendWait("{ESC}");
            Sleep(200);
            SendKeys.SendWait("^c"); 
            CheckClipboard(expected);

            Trace.WriteLine("Failed replace, via replace button");
            SendKeys.SendWait("{HOME}");
            TestForm.InvokeAsyncMenuItem("replaceToolStripMenuItem");
            findDialog = TestForm.WaitForPopup();
            SendKeys.SendWait("will not find%r");
            findDialog.ExpectingPopup("Find Error");
            findDialog.DismissPopUp("{ENTER}");
            TestForm.DismissPopUp("{ESC}");

            Trace.WriteLine("Check compound undo.");
            Undo();
            CheckOuterXml("<!-- This tests all element types -->");

            Sleep(1000);
            Save("out.xml");

            TestForm.Dispose();
            Sleep(2000);
            ResetFindOptions();
        }

        void ResetFindOptions() {
            string path = Environment.GetEnvironmentVariable("USERPROFILE") + "\\Local Settings\\Application Data\\Microsoft\\Xml Notepad\\XmlNotepad.settings";
            if (!File.Exists(path)) {
                return;
            }
            XmlDocument doc = new XmlDocument();
            doc.Load(path);

            RemoveNode(doc, "//SearchWholeWord");
            RemoveNode(doc, "//SearchRegex");
            RemoveNode(doc, "//SearchMatchCase");
            RemoveNode(doc, "//FindMode");
            doc.Save(path);
        }

        void ClearSchemaCache() {
            string path = Environment.GetEnvironmentVariable("USERPROFILE") + "\\Local Settings\\Application Data\\Microsoft\\Xml Notepad\\XmlNotepad.settings";
            if (!File.Exists(path)) {
                return;
            }
            XmlDocument doc = new XmlDocument();
            doc.Load(path);

            RemoveNodes(doc, "//Schema");
            doc.Save(path);
        }

        void RemoveNode(XmlDocument doc, string xpath) {
            XmlNode node = doc.SelectSingleNode(xpath);
            if (node != null) node.ParentNode.RemoveChild(node);
        }

        void RemoveNodes(XmlDocument doc, string xpath) {
            foreach (XmlNode node in doc.SelectSingleNode(xpath)) {
                node.ParentNode.RemoveChild(node);
            }
        }

        [TestMethod]
        public void TestToolbarAndContextMenus() {
            Trace.WriteLine("TestToolbarAndContextMenus==========================================================");

            string testFile = TestDir + "UnitTests\\test1.xml";
            LaunchNotepad(testFile);

            Trace.WriteLine("Test toopstrip 'new' button");
            TestForm.InvokeAsyncMenuItem("toolStripButtonNew");

            Trace.WriteLine("test recent files menu");
            SendKeys.SendWait("%f");
            Sleep(500);
            SendKeys.SendWait("f");
            Sleep(500);
            SendKeys.SendWait("{ENTER}");

            Sleep(1000);
            Trace.WriteLine("Test toolstrip button open");
            TestForm.InvokeAsyncMenuItem("toolStripButtonOpen");
            Sleep(500);
            SendKeys.SendWait(testFile + "{ENTER}");
            Sleep(500);
            SendKeys.SendWait("^IRoot");

            Trace.WriteLine("Bring up context menu");
            SendKeys.SendWait("^ ");
            Sleep(500);
            SendKeys.SendWait("{UP}{ENTER}"); // collapse
            // Bring up context menu
            SendKeys.SendWait("^ ");
            Sleep(500);
            SendKeys.SendWait("{UP}{UP}{ENTER}"); // expand

            Trace.WriteLine("Test toolstrip copy, cut, undo, delete");
            SendKeys.SendWait("{UP}");
            Clipboard.SetText("error");
            TestForm.InvokeMenuItem("toolStripButtonCopy");
            CheckClipboard("<?pi at root level?>");
            Clipboard.SetText("error");
            TestForm.InvokeMenuItem("toolStripButtonCut");
            CheckClipboard("<?pi at root level?>");
            TestForm.InvokeMenuItem("toolStripButtonUndo");
            TestForm.InvokeMenuItem("toolStripButtonDelete");
            SendKeys.SendWait("{UP}");

            Trace.WriteLine("test toolstrip paste, undo, redo");
            Clipboard.SetText("<?pi at root level?>");
            TestForm.InvokeMenuItem("toolStripButtonPaste");
            TestForm.InvokeMenuItem("toolStripButtonUndo");
            TestForm.InvokeMenuItem("toolStripButtonRedo");
            CheckNodeValue("pi");

            Trace.WriteLine("Test nudge commands");
            Sleep(1000);
            TestForm.InvokeMenuItem("toolStripButtonNudgeDown");
            TestForm.InvokeMenuItem("toolStripButtonNudgeRight");
            TestForm.InvokeMenuItem("toolStripButtonNudgeUp");
            TestForm.InvokeMenuItem("toolStripButtonNudgeLeft");

            Sleep(1000);
            Trace.WriteLine("context menu item - insert comment before");
            //bugbug: context menu items are not accessible?
            //TestForm.InvokeMenuItem("ctxCommentBeforeToolStripMenuItem");
            SendKeys.SendWait("^ mb");
            Sleep(200);
            SendKeys.SendWait("it is finished");
            //bugbug: TestForm.InvokeMenuItem("ctxPIBeforeToolStripMenuItem");
            SendKeys.SendWait("^ ob");
            Sleep(200);            
            SendKeys.SendWait("page{TAB}break{ENTER}");

            Save("out.xml");
            Sleep(1000);

            Trace.WriteLine("Test toolStripButtonSave 'save'");
            TestForm.InvokeMenuItem("toolStripButtonSave");            

            this.SaveAndCompare("out.xml", "test5.xml");
        }

        [TestMethod]
        public void TestNudge() {
            Trace.WriteLine("TestNudge==========================================================");
            string testFile = TestDir + "UnitTests\\test1.xml";
            LaunchNotepad(testFile);

            // better test when things are expanded
            TestForm.InvokeMenuItem("collapseAllToolStripMenuItem");
 
            // better test when things are expanded
            TestForm.InvokeMenuItem("expandAllToolStripMenuItem");
            
            int cmds = 0;
            SendKeys.SendWait("^I#"); // select first comment
            TestForm.InvokeMenuItem("downToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("downToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;


            // test nudge attr ({DOWN} resets type-to-find).
            Sleep(1000);
            SendKeys.SendWait("{DOWN}^Iid"); // select first attribute
            
            TestForm.InvokeMenuItem("downToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;

            // test nudge element .
            Sleep(1000);
            SendKeys.SendWait("{DOWN}^IEmp");
            
            TestForm.InvokeMenuItem("downToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("rightToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("leftToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;

            // test nudge pi
            Sleep(1000);
            SendKeys.SendWait("{DOWN}^Ipi"); // select next pi
            
            TestForm.InvokeMenuItem("leftToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("rightToolStripMenuItem");
            cmds++;
            TestForm.InvokeMenuItem("upToolStripMenuItem");
            cmds++;

            // Make sure MoveNode is undoable!
            UndoRedo(cmds);

            this.SaveAndCompare("out.xml", "test3.xml");
        }

        [TestMethod]
        public void TestDragDrop() {
            Trace.WriteLine("TestDragDrop==========================================================");
            this.LaunchNotepad();

            Rectangle treeBounds = this.XmlTreeView.Bounds;

            Trace.WriteLine("OpenFileDialog");
            TestForm.InvokeAsyncMenuItem("openToolStripMenuItem");
            Window openDialog = TestForm.WaitForPopup();
            Trace.WriteLine("Opening '" + TestDir + "UnitTests'");
            SendKeys.SendWait(TestDir + "UnitTests{ENTER}");
            Sleep(1000);

            // Drag/drop from open file dialog into xml notepad client area.
            bool found = true;
            SystemAccessible acc = (SystemAccessible)openDialog.AccessibleObject;
            Point drop = GetDropSpot(openDialog, treeBounds);
            Trace.WriteLine("Drop spot = " + drop.ToString());

            Rectangle obounds = acc.Bounds;            
            Point center = new Point(obounds.Left + (2 * obounds.Width / 3), obounds.Top + (obounds.Height/2));
            SystemAccessible list = SystemAccessible.AccessibleWindowObjectAt(center);
            if (list != null) {
                //Sleep(1000);
                AccessibleObject item = list.FindChild("test1.xml", new FileComparer());
                if (item != null) {
                    Rectangle ibounds = item.Bounds;
                    Point iloc = new Point(ibounds.Left + 10, ibounds.Top + 10);
                    Trace.WriteLine("Dragging from " + iloc.ToString());
                    Mouse.MouseDragDrop(iloc, drop, 5, MouseButtons.Left);
                    Sleep(1000);
                    SendKeys.SendWait("{ESC}");
                    found = true;
                }
            }
            if (!found) {
                SendKeys.SendWait("+{TAB}test1"); 
                SendKeys.SendWait("{ENTER}");
                Sleep(2000);
            }
            
            // need bigger window to test drag/drop
            TestForm.SetWindowSize(800, 600);
            
            TestForm.InvokeMenuItem("collapseAllToolStripMenuItem");
            TestForm.InvokeMenuItem("expandAllToolStripMenuItem");

            // Test mouse wheel
            AccessibleObject tree = this.TreeView;
            CheckProperties(tree);
            
            SendKeys.SendWait("{HOME}");
            Cursor.Position = Center(tree.Bounds);
            Sleep(500); // wait for focus to kick in before sending mouse events.
            Mouse.MouseWheel(-120 * 15); // first one doesn't get thru for some reason!
            Sleep(500);
            Mouse.MouseWheel(120 * 15);
            Sleep(500);
            
            // Test navigation keys
            SendKeys.SendWait("{HOME}");
            CheckNodeValue("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            SendKeys.SendWait("{END}");
            CheckNodeValue("<!--last comment-->");
            SendKeys.SendWait("{PGUP}{PGDN}{UP}{UP}");
            CheckNodeValue("Name");            
            
            // Get AccessibleObject to selected node in the tree.
            AccessibleObject ntv = this.NodeTextView;
            CheckProperties(ntv);
            // mouse down in node text view
            AccessibleObject node = ntv.GetFocused();
            AccessibleObject sel = ntv.GetSelected();
            node = node.Navigate(AccessibleNavigation.Next); // Office node.
            CheckNodeName(node, "Office");
            Rectangle bounds = node.Bounds;
            Mouse.MouseClick(Center(bounds), MouseButtons.Left);

            // test edit of node value using AccessibilityObject
            string office = "35/1682";
            string oldValue = node.Value;
            node.Value = office;
            CheckNodeValue(office);  // confirm via copy operation
                        
            node = tree.GetFocused();
            if (node == null) {
                throw new ApplicationException("Selected node not found");
            }
            CheckProperties(node);
            CheckNodeName(node, "Office");
            node.Select(AccessibleSelection.AddSelection);

            // test edit of node name using accessibility.
            node.Name = "MyOffice";            
            CheckNodeValue("MyOffice");  // confirm via copy operation

            // Test that "right arrow" moves over to the nodeTextView.
            SendKeys.SendWait("{RIGHT}{DOWN}{RIGHT}");
            CheckNodeValue("35/1682");  // confirm via copy operation

            Undo(); // make sure we can undo node name change!
            Undo(); // make sure we can undo node value change (while #text is expanded)!
            SendKeys.SendWait("{LEFT}");
            CheckNodeValue("Office");

            Trace.WriteLine("Select the 'Country' node.");
            bounds = node.Bounds;
            Trace.WriteLine(bounds.ToString());
            int itemHeight = bounds.Height;
            Point pt = Center(bounds);
            pt.Y -= (itemHeight * 2);

            // Test mouse down in tree view;
            Mouse.MouseClick(pt, MouseButtons.Left);
            Sleep(200);
            node = tree.GetFocused();
            CheckNodeName(node, "Country");

            Trace.WriteLine("Drag/drop country up 3 items");
            Sleep(1000); // avoid double click by delaying next click

            Point endPt = new Point(pt.X, pt.Y - (int)(3 * itemHeight));
            // Drag the node up three slots.
            Mouse.MouseDragDrop(pt, endPt, 5, MouseButtons.Left);

            Sleep(200); 
            
            node = tree.GetFocused();
            CheckNodeName(node, "Country");

            // Drag/drop to auto scroll, then leave the window and drop it on desktop
            Rectangle formBounds = TestForm.GetScreenBounds();
            Mouse.MouseDown(endPt, MouseButtons.Left);
            // Autoscroll
            Point treeTop = TopCenter(tree.Bounds, 2);

            Trace.WriteLine("--- Drag to top of tree view ---"); 
            Mouse.MouseDragTo(endPt, treeTop, 5, MouseButtons.Left);
            Sleep(1000); // autoscroll time.
            // Drag out of tree view.
            Point titleBar = TopCenter(formBounds, 20);
            Trace.WriteLine("--- Drag to titlebar ---");
            Mouse.MouseDragTo(treeTop, titleBar, 10, MouseButtons.Left);
            Sleep(1000); // should now have 'no drop icon'.
            Mouse.MouseUp(titleBar, MouseButtons.Left);            

            // code coverage on expand/collapse.
            SendKeys.SendWait("^IOffice");
            node.DoDefaultAction();
            Sleep(500);
            SendKeys.SendWait("{LEFT}");
            Sleep(500);
            SendKeys.SendWait("{RIGHT}");

            Sleep(1000);
            Trace.WriteLine("Test task list resizers");
            AccessibleObject resizer = TestForm.FindDescendant("TaskResizer");
            Trace.WriteLine(resizer.Parent.Name);
            bounds = resizer.Bounds;
            Point mid = Center(bounds);
            // Drag the resizer up a few pixels.
            Mouse.MouseDragDrop(mid, new Point(mid.X, mid.Y - 15), 2, MouseButtons.Left);

            Trace.WriteLine("Test tree view resizer");
            resizer = TestForm.FindDescendant("XmlTreeResizer");
            Trace.WriteLine(resizer.Parent.Name);
            bounds = resizer.Bounds;
            mid = Center(bounds);
            // Drag the resizer up a few pixels.
            Mouse.MouseDragDrop(mid, new Point(mid.X + 15, mid.Y), 2, MouseButtons.Left);

            this.SaveAndCompare("out.xml", "test4.xml");
        }

        Point GetDropSpot(Window w, Rectangle target) {
            AccessibleObject acc = w.AccessibleObject;
            Rectangle source = acc.Bounds;
            source.Inflate(20, 20); // add extra margin
            if (source.Contains(target)) {
                // Source window is completely occluding the target window, so we need to move it!
                Point from = new Point(source.Left + (source.Width/2), source.Top + 5);
                int amount = target.Left - source.Left + 100;
                Point end = new Point(from.X + amount, from.Y);
                // Move window to the right.
                Mouse.MouseDown(from, MouseButtons.Left);
                Mouse.MouseDragDrop(from, end, 5, MouseButtons.Left);
                
                source = acc.Bounds;
            }
            if (source.Left > target.Left) {
                // pick a spot along the left margin
                return new Point((target.Left + source.Left) / 2, (target.Top + target.Bottom) / 2);
            } else if (source.Right < target.Right) {
                // pick a spot along the right margin
                return new Point((target.Right + source.Right) / 2, (target.Top + target.Bottom) / 2);
            } else if (source.Top > target.Top) {
                // top margin
                return new Point((target.Right + target.Left) / 2, (source.Top + target.Top) / 2);
            } else if (source.Bottom < target.Bottom) {
                // bottom margin
                return new Point((target.Right + target.Left) / 2, (source.Bottom + target.Bottom) / 2);
            }

            // Then MOVE the window so it's not in the way!
            w.SetWindowPosition(target.Right, source.Top);
            Sleep(1000);
            source = acc.Bounds;
            return new Point((target.Left + source.Left) / 2, (target.Top + target.Bottom) / 2);
        }

        [TestMethod]
        public void TestInternetExplorerDragDrop() {
            Trace.WriteLine("TestInternetExplorerDragDrop==========================================================");

            this.LaunchNotepad();

            Rectangle treeBounds = this.XmlTreeView.Bounds;

            Trace.WriteLine("Launching IE");
            using (Window w = LaunchIE("http://msdn.microsoft.com/rss.xml")) {

                Sleep(5000); // give it plenty of time to settle down!.

                SystemAccessible ieframe = (SystemAccessible)w.AccessibleObject;
                SystemAccessible navbar = ieframe.FindChild("Navigation Bar");
                SystemAccessible addressBar = navbar.FindChild("Navigation Bar");

                Point dropPoint = GetDropSpot(w, treeBounds);
                Sleep(500);
                SystemAccessible textBox = addressBar.FirstChild.FirstChild.FirstChild.NextSibling;
                Rectangle bounds = textBox.Bounds;
                Point dragPoint = new Point(bounds.Left + 10, bounds.Top + 10);

                Trace.WriteLine("Dragging from: " + dragPoint.ToString());
                Trace.WriteLine("Drop spot: " + dropPoint.ToString());
                Sleep(500);
                Mouse.MouseDragDrop(dragPoint, dropPoint, 10, MouseButtons.Left);
                Sleep(500);

                Sleep(1000); 
                TestForm.Activate();
                Trace.WriteLine("Wait for rss to be loaded");
                this.WaitForText("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                SendKeys.SendWait("^Irss");
                this.CheckNodeName("rss");

                string outFile = Save("out.xml");
                Sleep(1000);

                XmlDocument doc = new XmlDocument();
                doc.Load(outFile);
                if (doc.DocumentElement.LocalName != "rss") {
                    throw new ApplicationException("Expected rss in UnitTests\\out.xml");
                }
            }

            return;
        }

        [TestMethod]
        public void TestAccessibility() {

            Trace.WriteLine("TestAccessibility==========================================================");
            string testFile = TestDir + "UnitTests\\test1.xml";
            this.LaunchNotepad(testFile);
            Sleep(1000);

            // Get AccessibleObject to selected node in the tree.
            AccessibleObject tree = this.TreeView;
            AccessibleObject root = tree.GetChild(4);
            
            // employee
            AccessibleObject emp = root.GetChild(8);
            emp.Select(AccessibleSelection.TakeSelection);
            CheckNodeName(emp, "Employee");
            Trace.Assert(emp.Name == emp.GetFocused().Name);
            Trace.Assert(emp.Name == root.GetFocused().Name);
            AccessibleObject node = emp.Navigate(AccessibleNavigation.FirstChild); // should expand

            // Test accesibility navigation!
            node = node.Parent;
            CheckNodeName(node, "Employee");
            node = node.Navigate(AccessibleNavigation.FirstChild);
            node = node.Navigate(AccessibleNavigation.Next);
            node.Select(AccessibleSelection.AddSelection);
            CheckNodeName(node, "id");
            node = node.Navigate(AccessibleNavigation.Right); // over to node text view!
            node.Select(AccessibleSelection.AddSelection);
            CheckProperties(node);
            CheckNodeValue(node, "46613");
            node = node.Navigate(AccessibleNavigation.Down);
            node.Select(AccessibleSelection.AddSelection);
            CheckNodeValue(node, "Architect");
            node = node.Navigate(AccessibleNavigation.Left);
            node.Select(AccessibleSelection.AddSelection);
            CheckNodeName(node, "title");

            node = node.Navigate(AccessibleNavigation.Right); // over to node text view!
            node = node.Navigate(AccessibleNavigation.Left);
            node = node.Navigate(AccessibleNavigation.Up);
            node = node.Navigate(AccessibleNavigation.Previous);
            CheckNodeName(node, "xmlns");
            node = node.Navigate(AccessibleNavigation.Down);
            node = node.Navigate(AccessibleNavigation.Next);
            CheckNodeName(node, "title");
            
            // Test TAB and SHIFT-TAB navigation.
            SendKeys.SendWait("{TAB}");
            Sleep(200);
            CheckNodeValue("Architect");
            SendKeys.SendWait("{TAB}");
            Sleep(200);
            CheckNodeValue("Name");
            SendKeys.SendWait("+{TAB}");
            Sleep(200);
            CheckNodeValue("Architect");
            SendKeys.SendWait("+{TAB}");
            Sleep(200);
            CheckNodeValue("title");

            // change node value!
            node.Value = "foo";
            SendKeys.SendWait("{TAB}");
            Sleep(200);
            CheckNodeValue("foo");
            SendKeys.SendWait("+{TAB}{LEFT}"); // back to Employee

            // hit test Employee node.
            Point p = Center(emp.Bounds);
            node = node.HitTest(p.X, p.Y);
            Trace.Assert(node.Name == emp.Name);

            emp.Select(AccessibleSelection.RemoveSelection);
            emp.Select(AccessibleSelection.TakeSelection);
            Trace.Assert(root == emp.Parent);
            AccessibleObject parent = root.Parent;
            string name = parent.Name;
            Trace.Assert(name == "TreeView");

            // default action on tree is toggle!
            tree.DoDefaultAction();
            Sleep(500);

            // state on invisible nodes.
            Trace.Assert((node.State & AccessibleStates.Invisible) == 0);

            // get last child of tree
            AccessibleObject cset = tree.Navigate(AccessibleNavigation.LastChild);
            CheckNodeName(cset, "CompletionSet");

            // select tree
            tree.Select(AccessibleSelection.TakeFocus);

            // Get AccessibleObject on node text view.
            AccessibleObject ntv = NodeTextView;
            ntv.Select(AccessibleSelection.TakeFocus);
            AccessibleObject first = ntv.Navigate(AccessibleNavigation.FirstChild);
            first.Select(AccessibleSelection.TakeFocus);
            ntv.DoDefaultAction(); // enter edit mode
            SendKeys.SendWait("{ESC}");

            // hit test on node text
            AccessibleObject pivalue = ntv.GetChild(3);
            CheckNodeValue(pivalue, "at root level");
            p = Center(pivalue.Bounds);
            Trace.Assert(ntv.HitTest(p.X, p.Y).Name == pivalue.Name);

            // Navigate to last child
            cset = ntv.Navigate(AccessibleNavigation.LastChild);
            CheckNodeName(cset, "CompletionSet");

            AccessibleObject next = first.Navigate(AccessibleNavigation.Next);
            CheckNodeValue(next, " This tests all element types ");
            Trace.Assert((next.State & AccessibleStates.Focused) == 0);

            next = next.Navigate(AccessibleNavigation.Next); // pi
            next = next.Navigate(AccessibleNavigation.Next); // root
            next.DoDefaultAction(); // toggle
            next.DoDefaultAction(); // toggle

            AccessibleObject ev = next.GetChild(8); // Employee
            ev.Select(AccessibleSelection.TakeSelection);
            Trace.Assert(ev.GetFocused().Name == ev.Name);
            SendKeys.SendWait("{RIGHT}"); // expand employee.

            node = ev.Navigate(AccessibleNavigation.FirstChild);
            CheckNodeValue(node, "http://www.hr.org");
            node = ev.Navigate(AccessibleNavigation.LastChild);
            CheckNodeName(node, "Office");
            node = node.Navigate(AccessibleNavigation.Up);
            CheckNodeName(node, "Country");
            node.Select(AccessibleSelection.RemoveSelection);
            node.Select(AccessibleSelection.TakeSelection);

            node.Name = "foo";
            CheckNodeName(node, "foo");
            Undo();
            CheckNodeName(node, "Country");

            Save("out.xml");
        }
        
        [TestMethod]
        public void TestKeyboard() {
            Trace.WriteLine("TestKeyboard==========================================================");
            string testFile = TestDir + "UnitTests\\emp.xml";
            string xsdFile = TestDir + "UnitTests\\emp.xsd";
            this.LaunchNotepad(testFile);

            Sleep(1000);

            Trace.WriteLine("Test goto definition on schemaLocation");
            SendKeys.SendWait("^Ixsi{F12}");
            Window popup = TestForm.ExpectingPopup("XML Notepad - " + xsdFile);
            TestForm.DismissPopUp("%{F4}");

            Trace.WriteLine("Test namespace intellisense, make sure emp.xsd namespace is in the list.");
            SendKeys.SendWait("{HOME}^Ixmlns{ESC}{TAB}{ENTER}");
            Sleep(250);
            SendKeys.SendWait("{END}{HOME}{ENTER}");
            Sleep(250);

            Trace.WriteLine("Test schemaLocation attribute.");
            SendKeys.SendWait("{LEFT}^Ixsi{DEL}");
            Sleep(1000); // let it validate without xsi location.
            Undo();
            Sleep(1000);

            Trace.WriteLine("Expand all");
            SendKeys.SendWait("{END}");
            SendKeys.SendWait("{MULTIPLY}"); // expandall.

            Trace.WriteLine("Goto next view");
            SendKeys.SendWait("{F6}"); // goto node text view
            Sleep(500);

            Trace.WriteLine("Create some validation errors");
            SendKeys.SendWait("{DOWN}^IRed");
            Sleep(200);
            SendKeys.SendWait("Red");
            Sleep(200);
            SendKeys.SendWait("{ENTER}{BACKSPACE}{ENTER}"); // delete "Redmond"
            Sleep(200);
            SendKeys.SendWait("{UP}^I98");
            Sleep(200);
            SendKeys.SendWait("{ENTER}{BACKSPACE}{ENTER}"); // delete "98052"

            Sleep(1000);  // give it a chance to validate and produce errors.

            Trace.WriteLine("Navigate errors");
            SendKeys.SendWait("{HOME}{F6}"); // Navigate to error list
            SendKeys.SendWait("{DOWN}{ENTER}"); // Select second error
            CheckNodeValue("Zip");

            Trace.WriteLine("Previous pane");
            SendKeys.SendWait("+{F6}"); // Navigate back to error list
            SendKeys.SendWait("{UP}{ENTER}"); // Select first error
            CheckNodeValue("City");            

            Trace.WriteLine("Next Error");
            SendKeys.SendWait("{F4}"); // next error
            CheckNodeValue("Zip");

            Trace.WriteLine("Collapse/Expand country");

            SendKeys.SendWait("{DOWN}^ICo{SUBTRACT}"); // collapse country
            Sleep(100); // just so we can watch it
            SendKeys.SendWait("{ADD}"); // re-expand country
            Sleep(100);

            Trace.WriteLine("Nudge commands");
            SendKeys.SendWait("^+{LEFT}"); // nudge country left
            Sleep(100);
            SendKeys.SendWait("^+{RIGHT}"); // nudge country right
            Sleep(100);
            SendKeys.SendWait("^+{UP}^+{UP}^+{UP}"); // nudge country back up to where it was
            Sleep(100);
            SendKeys.SendWait("^+{DOWN}"); // nudge country down
            Sleep(100);

            Trace.WriteLine("Fix errors");
            SendKeys.SendWait("{UP}{TAB}{F2}98052{ENTER}"); // add zip code back
            Sleep(200);
            SendKeys.SendWait("{UP}{ENTER}Redmond{ENTER}"); // add redmond back
            Sleep(300); // let it re-validate

            Trace.WriteLine("Test direct editing of names/values");
            SendKeys.SendWait("y{ENTER}");
            Sleep(100);
            SendKeys.SendWait("{LEFT}x{ENTER}");
            Sleep(100);
            SendKeys.SendWait("^c");

            CheckClipboard("<x xmlns=\"http://Employees\">y</x>");
            Undo();
            Undo();

            Sleep(1000);
            this.SaveAndCompare("out.xml", "emp.xml");
        }

        [TestMethod]
        public void TestMouse() {
            Trace.WriteLine("TestMouse==========================================================");
            string testFile = TestDir + "UnitTests\\emp.xml";
            this.LaunchNotepad(testFile);

            Sleep(1000);

            // Test mouse click on +/-.
            AccessibleObject tree = this.TreeView;
            AccessibleObject node = tree.Navigate(AccessibleNavigation.FirstChild);
            node = node.Navigate(AccessibleNavigation.LastChild);
            node.Select(AccessibleSelection.TakeSelection);

            Rectangle bounds = node.Bounds;
            TestHitTest(Center(bounds), tree, node);

            bool expanded = (node.State & AccessibleStates.Expanded) != 0;
            if (expanded){
                throw new ApplicationException(
                    string.Format("Did not expect node '{0}' to be expanded here", node.Name));
            }
            // minus tree indent and image size
            Point plusminus = new Point(bounds.Left - 30 - 16, (bounds.Top + bounds.Bottom) / 2);

            Mouse.MouseClick(plusminus, MouseButtons.Left);

            Sleep(500);

            bool expanded2 = (node.State & AccessibleStates.Expanded) != 0;
            if (!expanded2) {
                throw new ApplicationException("Node did not become expanded");
            }

            //mouse down edit of node name
            Mouse.MouseClick(Center(bounds), MouseButtons.Left);
            Sleep(1000); // give it enough time to kick into edit mode.

            CheckOuterXml("Employee");
            SendKeys.SendWait("{ESCAPE}");
            
            // code coverage on scrollbar interaction
            AccessibleObject vscroll = TestForm.FindDescendant("VScrollBar");
            bounds = vscroll.Bounds;

            Point downArrow = new Point((bounds.Left + bounds.Right) / 2, bounds.Bottom - (bounds.Width / 2));
            for (int i = 0; i < 10; i++) {
                Mouse.MouseClick(downArrow, MouseButtons.Left);
                Sleep(500);
            }

        }

        [TestMethod]
        public void TestUtilities() {
            Trace.WriteLine("TestUtilities==========================================================");
            // code coverage on hard to reach utility code.
            HLSColor hls = new HLSColor(Color.Red);
            Trace.WriteLine(hls.ToString());
            Trace.WriteLine(hls.Darker(0.5F).ToString());
            Trace.WriteLine(hls.Lighter(0.5F).ToString());
            Trace.WriteLine(hls == new HLSColor(Color.Red));
            Trace.WriteLine(hls.GetHashCode());

            // Test resource class.
            Type t = FormMain.ResourceType;
            foreach (PropertyInfo pi in t.GetProperties(BindingFlags.Static)) {
                if (pi.PropertyType == typeof(string)) {
                    string name = pi.Name;
                    object res = pi.GetValue(null, null);
                    if (res == null) {
                        throw new Exception("Unexpected null returned from property: " + name);
                    }
                    Trace.WriteLine(string.Format("{0}={1}", name, res.ToString()));
                }
            }

            // Test XmlIncludeReader
            string test = TestDir + "UnitTests\\includes\\index.xml";
            Uri baseUri = new Uri(test);
            XmlDocument doc = new XmlDocument();
            doc.Load(test);
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ProhibitDtd = false;
            foreach (XmlElement e in doc.SelectNodes("test/case")) {
                Uri input = new Uri(baseUri, e.GetAttribute("input"));
                Uri output = new Uri(baseUri, e.GetAttribute("results"));
                using (XmlIncludeReader r = XmlIncludeReader.CreateIncludeReader(input.LocalPath, settings)) {
                    CompareResults(ReadNodes(r), output.LocalPath);
                }
            }
        }

        [TestMethod]
        public void TestNoBorderTabControl() {
            Form f = new Form();
            f.Size = new Size(400, 400);

            NoBorderTabControl tabs = new NoBorderTabControl();
            NoBorderTabPage page1 = new NoBorderTabPage();
            page1.Text = "Apple";
            CheckBox c = new CheckBox();
            c.Text = "this is a checkbox";
            c.Location = new Point(10, 10);
            page1.Controls.Add(c);
            tabs.TabPages.Add(page1);
            NoBorderTabPage page2 = new NoBorderTabPage();
            page2.Text = "Orange";
            RadioButton b1 = new RadioButton();
            b1.Text = "test1";
            b1.Location = new Point(10, 10); 
            page2.Controls.Add(b1);
            RadioButton b2 = new RadioButton();
            b2.Text = "test2";
            b2.Location = new Point(10, 30);
            page2.Controls.Add(b2);
            tabs.TabPages.Insert(0, page2);
            tabs.Dock = DockStyle.Fill;
            f.Controls.Add(tabs);

            f.Show();
            
            Sleep(1000);

            tabs.TabPages.Remove(page1);
            tabs.TabPages.Remove(page2);
            tabs.TabPages.Add(page1);
            tabs.TabPages.Add(page2);

            
            Sleep(1000);

            tabs.TabPages.Clear();
            tabs.TabPages.Add(page1);

            
            Sleep(1000);

            Trace.Assert(tabs.TabPages.Contains(page1));
            Trace.Assert(!tabs.TabPages.Contains(page2));
            tabs.TabPages.Insert(0, page2);

            
            Sleep(1000);

            int i = tabs.TabPages.IndexOf(page1);
            Trace.Assert(i == 1);

            i = tabs.TabPages.IndexOf(page2);
            Trace.Assert(i == 0);

            Trace.Assert(!tabs.TabPages.IsFixedSize);
            Trace.Assert(!tabs.TabPages.IsReadOnly);
            Trace.Assert(!tabs.TabPages.IsSynchronized);

            tabs.TabPages.Remove(page1);
            tabs.TabPages.RemoveAt(0);
            
            Sleep(1000);

            tabs.TabPages[0] = page1;
            tabs.TabPages[1] = page2;
            
            Sleep(1000);

            NoBorderTabPage[] a = new NoBorderTabPage[tabs.TabPages.Count];
            tabs.TabPages.CopyTo(a, 0);
            Trace.Assert(a[0] == page1);
            Trace.Assert(a[1] == page2);
            f.Close();
        }

        [TestMethod]
        public void TestInclude() {
            Trace.WriteLine("TestInclude==========================================================");
            string nonexist = TestDir + "UnitTests\\Includes\\nonexist.xml";
            WipeFile(nonexist);
            try {                

                string testFile = TestDir + "UnitTests\\Includes\\i1.xml";
                this.LaunchNotepad(testFile);

                SendKeys.SendWait("^Iinclude");
                TestForm.InvokeMenuItem("gotoDefinitionToolStripMenuItem");
                TestForm.WaitForNewWindow();
                TestForm.DismissPopUp("%{F4}");
                
                TestForm.InvokeMenuItem("expandXIncludesToolStripMenuItem");
                this.SaveAndCompare("Includes\\out.xml", "Includes\\r1.xml");

                Trace.WriteLine("Test F12 on non-existant include");
                testFile = TestDir + "UnitTests\\Includes\\i3.xml";
                TestForm.InvokeAsyncMenuItem("openToolStripMenuItem");
                TestForm.WaitForPopup();
                SendKeys.SendWait(testFile + "{ENTER}");
               
                SendKeys.SendWait("^Ix:{F12}");
                TestForm.WaitForPopup();
                
                SendKeys.SendWait("{ENTER}");
                // Should create new file and open it!
                Sleep(1000);
                TestForm.WaitForPopup();
                SendKeys.SendWait("%{F4}");

                Sleep(2000);
                if (!File.Exists(nonexist)) {
                    throw new ApplicationException("File should now exist!");
                }
            } finally {
                WipeFile(nonexist);
            }
        }

        [TestMethod]
        public void TestUnicode() {

            ClearSchemaCache();

            Trace.WriteLine("TestUnicode==========================================================");
            string testFile = TestDir + "UnitTests\\unicode.xml";
            this.LaunchNotepad(testFile);

            string outFile = TestDir + "UnitTests\\out.xml";
            WipeFile(outFile);

            TestForm.InvokeAsyncMenuItem("exportErrorsToolStripMenuItem");
            
            TestForm.WaitForPopup();
            TestForm.DismissPopUp(outFile + "{ENTER}");
            
            string expectedFile = TestDir + "UnitTests\\errors.xml";
            CompareResults(ReadNodes(expectedFile), outFile);

        }

        [TestMethod]
        public void TestChangeTo() {
            Trace.WriteLine("TestChangeTo==========================================================");
            string testFile = TestDir + "UnitTests\\test8.xml";
            this.LaunchNotepad(testFile);
            Sleep(1000);

            TestForm.InvokeMenuItem("expandAllToolStripMenuItem");
            Sleep(1000);

            SendKeys.SendWait("{DOWN}^ICard{ESC}");
            Sleep(200);
            SendKeys.SendWait("^c");
            string expected = Clipboard.GetText(); // save expected text

            Trace.WriteLine("Change element to attribute");
            SendKeys.SendWait("{DOWN}{DOWN}");
            TestForm.InvokeMenuItem("changeToAttributeToolStripMenuItem1");
            Sleep(1000);
            SendKeys.SendWait("{LEFT}");
            Sleep(200);
            SendKeys.SendWait("^c");
            this.CheckClipboard("<Card bar=\"2\"><foo>1</foo><end>3</end></Card>");
            Trace.WriteLine("Make undo inserts element in the right place.");
            this.Undo();
            SendKeys.SendWait("{LEFT}");
            Sleep(200);
            SendKeys.SendWait("^c");
            this.CheckClipboard(expected);

            Sleep(2000);
            SendKeys.SendWait("{DOWN}^IName{ESC}");
            TestForm.InvokeMenuItem("changeToCDATAToolStripMenuItem1");
            this.CheckOuterXml("<![CDATA[<Name First=\"Chris\" Last=\"Lovett\">/[A CDATA block]/</Name>]]>");

            Trace.WriteLine("Change element to Comment (with nested comments).");
            SendKeys.SendWait("{LEFT}");
            TestForm.InvokeMenuItem("changeToCommentToolStripMenuItem1");
            this.CheckOuterXml("<!--<Contact>/*inner comment*/<![CDATA[<Name First=\"Chris\" Last=\"Lovett\">/[A CDATA block]/</Name>]]></Contact>-->");

            Trace.WriteLine("Change comment back to element (with nested comments!)");
            TestForm.InvokeMenuItem("changeToElementToolStripMenuItem1");
            this.CheckOuterXml("<Contact><!--inner comment--><![CDATA[<Name First=\"Chris\" Last=\"Lovett\">/[A CDATA block]/</Name>]]></Contact>");

            Trace.WriteLine("Change CDATA back to element (with nested CDATA!)");
            SendKeys.SendWait("{END}");
            TestForm.InvokeMenuItem("changeToElementToolStripMenuItem1");//bugbug: "changeToElementContextMenuItem"); // these are not accessible
            this.CheckOuterXml("<Name First=\"Chris\" Last=\"Lovett\"><![CDATA[A CDATA block]]></Name>");

            Trace.WriteLine("Make sure this is all undoable.");
            this.Undo();
            this.Undo();
            this.Undo();
            this.Undo();

            Trace.WriteLine("Now file should be identical to original");            
            SaveAndCompare("out.xml", "test8.xml");

            Trace.WriteLine("Change attribute to element");
            SendKeys.SendWait("^Iid{ESC}");
            TestForm.InvokeMenuItem("changeToElementToolStripMenuItem1");
            this.CheckOuterXml("<id>55</id>");

            Trace.WriteLine("Change element to attribute");
            TestForm.InvokeMenuItem("changeToAttributeToolStripMenuItem1");//changeToAttributeContextMenuItem");
            this.CheckOuterXml("id=\"55\"");

            this.Undo();
            this.Undo();

            Trace.WriteLine("Change attribute to PI");
            TestForm.InvokeMenuItem("changeToProcessingInstructionToolStripMenuItem");
            this.CheckOuterXml("<?id 55?>");

            Trace.WriteLine("Change PI to element");
            TestForm.InvokeMenuItem("changeToElementToolStripMenuItem1");
            this.CheckOuterXml("<id>55</id>");

            Trace.WriteLine("Change element to text");
            TestForm.InvokeMenuItem("changeToTextToolStripMenuItem1");
            this.CheckOuterXml("&lt;id&gt;55&lt;/id&gt;");

            Trace.WriteLine("Change text to comment");
            TestForm.InvokeMenuItem("changeToCommentToolStripMenuItem1");//changeToCommentContextMenuItem");
            this.CheckOuterXml("<!--<id>55</id>-->");

            Trace.WriteLine("Change comment to CDATA");
            TestForm.InvokeMenuItem("changeToCDATAToolStripMenuItem1");//changeToCDATAContextMenuItem");
            this.CheckOuterXml("<![CDATA[<id>55</id>]]>");

            Trace.WriteLine("Change CDATA to Attribute");
            TestForm.InvokeMenuItem("changeToAttributeToolStripMenuItem1");
            this.CheckOuterXml("id=\"55\"");

            this.Undo();
            this.Undo();
            this.Undo();
            this.Undo();
            this.Undo();
            this.Undo();

            SendKeys.SendWait("{END}");
            Trace.WriteLine("Change CDATA to Comment");
            TestForm.InvokeMenuItem("changeToCommentToolStripMenuItem1");
            this.CheckOuterXml("<!--A CDATA block-->");

            Trace.WriteLine("Change Comment to Text");
            TestForm.InvokeMenuItem("changeToTextToolStripMenuItem1");
            this.CheckOuterXml("A CDATA block");

            Sleep(1000);
            Trace.WriteLine("Change text to PI");
            TestForm.InvokeMenuItem("changeToProcessingInstructionToolStripMenuItem");//changeToProcessingInstructionContextMenuItem");
            this.CheckOuterXml("<?pi A CDATA block?>");

            Trace.WriteLine("Change PI to attribute");
            TestForm.InvokeMenuItem("changeToAttributeToolStripMenuItem1");
            this.CheckOuterXml("pi=\"A CDATA block\"");

            Trace.WriteLine("Change attribute to comment");
            TestForm.InvokeMenuItem("changeToCommentToolStripMenuItem1");
            this.CheckOuterXml("<!--pi=\"A CDATA block\"-->");

            Trace.WriteLine("Change comment to PI");
            TestForm.InvokeMenuItem("changeToProcessingInstructionToolStripMenuItem");
            this.CheckOuterXml("<?pi A CDATA block?>");

            Trace.WriteLine("Change PI to comment");
            TestForm.InvokeMenuItem("changeToCommentToolStripMenuItem1");
            this.CheckOuterXml("<!--<?pi A CDATA block?>-->");

            this.Undo();

            Trace.WriteLine("Change PI to CDATA");
            TestForm.InvokeMenuItem("changeToCDATAToolStripMenuItem1");
            this.CheckOuterXml("<![CDATA[<?pi A CDATA block?>]]>");

            Trace.WriteLine("Change CDATA to Text");
            TestForm.InvokeMenuItem("changeToTextToolStripMenuItem1");
            this.CheckOuterXml("&lt;?pi A CDATA block?&gt;");

            Trace.WriteLine("Change Text to Element");
            TestForm.InvokeMenuItem("changeToElementToolStripMenuItem1");
            this.CheckOuterXml("<pi>A CDATA block</pi>");

            Trace.WriteLine("Change Element to PI");
            TestForm.InvokeMenuItem("changeToProcessingInstructionToolStripMenuItem");
            this.CheckOuterXml("<?pi A CDATA block?>");

            Trace.WriteLine("Change PI to Text");
            TestForm.InvokeMenuItem("changeToTextToolStripMenuItem1");
            this.CheckOuterXml("&lt;?pi A CDATA block?&gt;");

            Trace.WriteLine("Change Text to Attribute");
            TestForm.InvokeMenuItem("changeToAttributeToolStripMenuItem1");
            this.CheckOuterXml("pi=\"A CDATA block\"");

            Trace.WriteLine("Change Attribute to Text");
            TestForm.InvokeMenuItem("changeToTextToolStripMenuItem1");
            this.CheckOuterXml("pi=\"A CDATA block\"");

            Trace.WriteLine("Change Attribute to Text");
            TestForm.InvokeMenuItem("changeToCDATAToolStripMenuItem1");
            this.CheckOuterXml("<![CDATA[pi=\"A CDATA block\"]]>");

            Trace.WriteLine("Change CDATA to PI");
            TestForm.InvokeMenuItem("changeToProcessingInstructionToolStripMenuItem");
            this.CheckOuterXml("<?pi A CDATA block?>");

            Trace.WriteLine("Change PI to Element");
            TestForm.InvokeMenuItem("changeToElementToolStripMenuItem1");
            this.CheckOuterXml("<pi>A CDATA block</pi>");

            Trace.WriteLine("Change Element to PI");
            TestForm.InvokeMenuItem("changeToProcessingInstructionToolStripMenuItem");
            this.CheckOuterXml("<?pi A CDATA block?>");

            Trace.WriteLine("Change PI to Text");
            TestForm.InvokeMenuItem("changeToTextToolStripMenuItem1");
            this.CheckOuterXml("&lt;?pi A CDATA block?&gt;");

            Trace.WriteLine("Change Text to PI");
            TestForm.InvokeMenuItem("changeToProcessingInstructionToolStripMenuItem");
            this.CheckOuterXml("<?pi A CDATA block?>");

            this.Undo(19);

            Trace.WriteLine("Now file should be identical to original");
            this.SaveAndCompare("out.xml", "test8.xml");
        }

        //==================================================================================
        private void SaveAndCompare(string outname, string compareWith) {

            string outFile = Save(outname);

            string expectedFile = TestDir + "UnitTests\\" + compareWith;
            Sleep(1000);
            CompareResults(ReadNodes(expectedFile), outFile);
        }

        private string Save(string outname) {
            Trace.WriteLine("Save");            
            string outFile = TestDir + "UnitTests\\" + outname;
            DeleteFile(outFile);
            
            TestForm.InvokeAsyncMenuItem("saveAsToolStripMenuItem");
            TestForm.WaitForPopup();
            SendKeys.SendWait(outFile + "{ENTER}");
            return outFile;
        }


        void TestHitTest(Point pt, AccessibleObject parent, AccessibleObject expected) {
            AccessibleObject obj = parent.HitTest(pt.X, pt.Y);
            if (obj.Name != expected.Name) {
                throw new ApplicationException(
                    string.Format("Found node '{0}' at {1},{2} instead of node '{3}'",
                        obj.Name, pt.X.ToString(), pt.Y.ToString(), expected.Name)
                    );
            }
        }

        Point TopCenter(Rectangle bounds, int dy) {
            return new Point(bounds.Left + (bounds.Width / 2), bounds.Top + dy);
        }

        void FocusTreeView() {
            AccessibleObject acc = this.TreeView;
            AccessibleObject node = acc.GetFocused();
            if (node == null) {
                node = acc.Navigate(AccessibleNavigation.FirstChild);
            }
            node.Select(AccessibleSelection.TakeSelection);
        }

        void CheckNodeName(string expected) {
            AccessibleObject acc = this.TreeView;
            AccessibleObject node = acc.GetFocused();
            if (node == null) {
                throw new ApplicationException("No node selected in tre view!");
            }
            CheckNodeName(node, expected);
        }

        void CheckNodeName(AccessibleObject acc, string expected) {
            string name = acc.Name;
            if (name != expected) {
                throw new ApplicationException(string.Format("Expecting node name '{0}'", expected));
            }
            Trace.WriteLine("Name=" + name);
#if DEBUG
            Sleep(200); // so we can watch it!
#endif
        }

        void CheckNodeValue(AccessibleObject acc, string expected) {
            string value = acc.Value;
            if (value != expected) {
                throw new ApplicationException(string.Format("Expecting node value '{0}'", expected));
            }
            Trace.WriteLine("Value=" + value);
#if DEBUG
            Sleep(200); // so we can watch it!
#endif
        }

        void CheckProperties(AccessibleObject node) {
            // Get code coverage on the boring stuff.
            Trace.WriteLine("Name=" + node.Name);
            Trace.WriteLine("\tValue=" + node.Value);
            Trace.WriteLine("\tParent=" + node.Parent.Name);
            Trace.WriteLine("\tChildCount=" + node.GetChildCount());
            Trace.WriteLine("\tBounds=" + node.Bounds.ToString());
            Trace.WriteLine("\tDefaultAction=" + node.DefaultAction);
            Trace.WriteLine("\tDescription=" + node.Description);
            Trace.WriteLine("\tHelp=" + node.Help);
            Trace.WriteLine("\tKeyboardShortcut=" + node.KeyboardShortcut);
            Trace.WriteLine("\tRole=" + node.Role);
            Trace.WriteLine("\tState=" + node.State);
            string filename = null;
            Trace.WriteLine("\tHelpTopic=" + node.GetHelpTopic(out filename));
        }

        public override void CheckClipboard(string expected) {
            if (!Clipboard.ContainsText()) {
                throw new ApplicationException("clipboard does not contain any text!");
            }
            AssertNormalizedEqual(Clipboard.GetText(), expected);
        }

        public void AssertNormalizedEqual(string value, string expected) {
            expected = NormalizeNewLines(expected);
            string text = NormalizeNewLines(value);
            if (text != expected) {
                throw new ApplicationException(string.Format(@"clipboard '{0}' does not match expected value:
{1}", text, expected));
            }
        }

        public static string NormalizeNewLines(string text) {
            if (text == null) return null;
            StringBuilder sb = new StringBuilder();
            for (int i = 0, n = text.Length; i < n; i++) {
                char ch = text[i];
                if (ch == '\r') {
                    if (i + 1 < n && text[i + 1] == '\n')
                        i++;
                    sb.Append("\r\n");
                } else if (ch == '\n') {
                    sb.Append("\r\n");
                } else {
                    sb.Append(ch);
                }
            }
            return sb.ToString();
        }

    }
}
