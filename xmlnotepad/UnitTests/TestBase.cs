using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Threading;
using System.Xml;
using Accessibility;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Drawing;
using System.IO;

namespace UnitTests {
    public class TestBase {

        public class NodeInfo {
            XmlNodeType nt;
            string name;
            string value;
            public NodeInfo(XmlReader r) {
                this.nt = r.NodeType;
                this.name = r.Name;
                this.value = Normalize(r.Value);
            }
            public bool Equals(NodeInfo other) {
                return this.nt == other.nt && this.name == other.name &&
                    this.value == other.value;
            }

            string Normalize(string value) {
                // So text indented different still compares as the same.
                if (string.IsNullOrEmpty(value)) return null;
                StringBuilder sb = new StringBuilder();
                bool wasnewline = true; // collapse leading spaces
                for (int i = 0, n = value.Length; i < n; i++) {
                    char ch = value[i];
                    if (ch == '\r'){
                        if (i + 1 < n && value[i + 1] == '\n') {
                            i++;
                        }
                        sb.Append('\n');
                        wasnewline = true;
                    } else if (ch == '\n'){
                        sb.Append(ch);
                        wasnewline = true;
                    } else if (Char.IsWhiteSpace(ch)) {
                        if (!wasnewline) sb.Append(' ');
                    } else {
                        sb.Append(ch);
                        wasnewline = false;
                    }
                }
                return sb.ToString();
            }
        }

        Window main;

        public void Sleep(int ms) {
            Thread.Sleep(ms);
        }

        public Window LaunchApp(string exeFileName, string args) {
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = exeFileName;
            info.Arguments = args;

            Process p = new Process();
            p.StartInfo = info;
            if (!p.Start()) {
                string msg = "Error launching " + exeFileName;
                MessageBox.Show("Error Creating Process", msg, 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new Exception(msg);
            }
            if (p.HasExited) {
                throw new Exception(string.Format("Failed to launch '{0'}, exit code {1}", exeFileName, p.ExitCode.ToString()));
            }
            Window w = new Window(p);
            w.TestBase = this;
            if (this.main == null) {
                this.main = w;
            }
            return w;
        }


        public void CloseApp() {
            if (this.main != null) {
                this.main.Dispose();
                this.main = null;
            }
        }

        public Window TestForm {
            get { return main; }
        }
        
        public static Point Center(Rectangle bounds) {
            return new Point(bounds.Left + (bounds.Width / 2),
                bounds.Top + (bounds.Height / 2));
        }

        public void DeleteFile(string fname) {
            if (File.Exists(fname))
                File.Delete(fname);
        }

        public void CompareResults(List<NodeInfo> nodes, string outFile) {
            int pos = 0;
            XmlReader reader = XmlReader.Create(outFile);
            IXmlLineInfo li = (IXmlLineInfo)reader;
            using (reader) {
                while (reader.Read()) {
                    if (reader.NodeType == XmlNodeType.Whitespace ||
                        reader.NodeType == XmlNodeType.SignificantWhitespace ||
                        reader.NodeType == XmlNodeType.XmlDeclaration)
                        continue;

                    NodeInfo node = new NodeInfo(reader);
                    if (pos >= nodes.Count) {
                        throw new ApplicationException("Found too many nodes");
                    }
                    NodeInfo other = nodes[pos++];
                    if (!node.Equals(other)) {
                        throw new ApplicationException(
                            string.Format("Mismatching nodes at line {0},{1}",
                            li.LineNumber, li.LinePosition));
                    }
                }
            }
        }

        public List<NodeInfo> ReadNodes(string fileName) {
            XmlReader reader = XmlReader.Create(fileName);
            return ReadNodes(reader);
        }

        public List<NodeInfo> ReadNodes(XmlReader reader) {        
            List<NodeInfo> nodes = new List<NodeInfo>();
            using (reader) {
                while (reader.Read()) {
                    if (reader.NodeType == XmlNodeType.Whitespace ||
                        reader.NodeType == XmlNodeType.SignificantWhitespace ||
                        reader.NodeType == XmlNodeType.XmlDeclaration)
                        continue;

                    nodes.Add(new NodeInfo(reader));
                }
            }
            return nodes;
        }

        public virtual void CheckClipboard(string expected) {
            if (!Clipboard.ContainsText()) {
                throw new ApplicationException("clipboard does not contain any text!");
            }
            string text = Clipboard.GetText();
            if (text != expected) {
                throw new ApplicationException("clipboard does not match expected cut node");
            }
        }

        public Window LaunchIE(string args) {
            string ie = Environment.GetEnvironmentVariable("ProgramFiles") + "\\Internet Explorer\\iexplore.exe";
            return LaunchApp(ie, args);
        }



    }

    
}
