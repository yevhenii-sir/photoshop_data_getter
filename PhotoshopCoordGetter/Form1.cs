using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ntreev.Library.Psd;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.IO;
using System.Diagnostics;

namespace PhotoshopCoordGetter
{
    public partial class Form1 : Form
    {
        public string path_file = "";
        public string floder_to_save_image;
        PsdDocument document;
        Dictionary<string, IPsdLayer> PSDLayersDictionary = new Dictionary<string, IPsdLayer>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (File.Exists("previous.ini"))
            {
                using (StreamReader streamReader = new StreamReader("previous.ini"))
                {
                    string tempString;
                    while ((tempString = streamReader.ReadLine()) != null)
                    {
                        listBox1.Items.Add(tempString);
                    }
                }
            }

            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        public static string GetUniqueName(List<string> all_names, string name, int num)
        {
            if (all_names.Exists(str => str == (name + "_" + num)))
            {
                return GetUniqueName(all_names, name, num + 1);
            } else
            {
                return name + "_" + num;
            }
        }

        public TreeNode ShowLayers(IPsdLayer[] document, string name_tree, List<string> all_names)
        {
            TreeNode tempTree = new TreeNode(name_tree);

            foreach (var item in document)
            {
                if (item.Childs.Length > 0)
                {
                    tempTree.Nodes.Add(ShowLayers(item.Childs, item.Name, all_names));
                }
                else
                {
                    var name = item.Name.ToLower();
                    if (all_names.Exists(str => str == name)) 
                        name = GetUniqueName(all_names, name, 1);
                    all_names.Add(name);

                    tempTree.Nodes.Add(name);
                    PSDLayersDictionary[name] = item;
                }
            }

            return tempTree;
        }

        public void SaveList()
        {
            using (StreamWriter streamWriter = new StreamWriter(new FileStream("previous.ini", FileMode.Create)))
            {
                foreach (var text in listBox1.Items)
                {
                    streamWriter.WriteLine(text);
                }
            }
        }

        internal static IEnumerable<TreeNode> Descendants(TreeNodeCollection c)
        {
            foreach (var node in c.OfType<TreeNode>())
            {
                yield return node;

                foreach (var child in Descendants(node.Nodes))
                {
                    yield return child;
                }
            }
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Nodes.Count > 0)
            {
                for (var i = 0; i < e.Node.Nodes.Count; i++)
                {
                    e.Node.Nodes[i].Checked = e.Node.Checked;
                }
            }
        }

        public static void SaveImageToFile(BitmapSource file, String name, double scale)
        {
            file = new TransformedBitmap(file, new ScaleTransform(scale, scale));

            using (var fileStream = new FileStream(name + ".png", FileMode.Create))
            {
                BitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(file));
                encoder.Save(fileStream);
            }

            GC.Collect(0, GCCollectionMode.Forced);
            GC.WaitForPendingFinalizers();
        }

        public static BitmapSource GetBitmap(IImageSource imageSource)
        {
            if (imageSource.HasImage == false)
                return null;

            byte[] data = imageSource.MergeChannels();
            var channelCount = imageSource.Channels.Length;
            var pitch = imageSource.Width * imageSource.Channels.Length;
            var w = imageSource.Width;
            var h = imageSource.Height;

            var colors = new System.Windows.Media.Color[data.Length / channelCount];

            var k = 0;
            for (var y = h - 1; y >= 0; --y)
            {
                for (var x = 0; x < pitch; x += channelCount)
                {
                    var n = x + y * pitch;

                    var c = System.Windows.Media.Color.FromArgb(1, 1, 1, 1);
                    if (channelCount == 4)
                    {
                        c.B = data[n++];
                        c.G = data[n++];
                        c.R = data[n++];
                        c.A = (byte)System.Math.Round(data[n++] / 255f * imageSource.Opacity * 255f);
                    }
                    else
                    {
                        c.B = data[n++];
                        c.G = data[n++];
                        c.R = data[n++];
                        c.A = (byte)System.Math.Round(imageSource.Opacity * 255f);
                    }

                    colors[k++] = c;
                }
            }
            
            if (channelCount == 4)
                return BitmapSource.Create(imageSource.Width, imageSource.Height, 96, 96, 
                    PixelFormats.Bgra32, null,data, pitch);
            return BitmapSource.Create(imageSource.Width, imageSource.Height, 96, 96, 
                PixelFormats.Bgr24, null, data, pitch);
        }

        private (double x, double y) GetCoordinateWithOrigin(IPsdLayer item)
        {
            double x_center = item.Left + (item.Width / 2);
            double y_center = item.Top + (item.Height / 2);

            if (radio_top_left.Checked) return (item.Left, item.Top);
            if (radio_top_center.Checked) return (x_center, item.Top);
            if (radio_top_right.Checked) return (item.Right, item.Top);

            if (radio_middle_left.Checked) return (item.Left, y_center);
            if (radio_middle_center.Checked) return (x_center, y_center);
            if (radio_middle_right.Checked) return (item.Right, y_center);
            
            if (radio_bottom_left.Checked) return (item.Left, item.Bottom);
            if (radio_bottom_center.Checked) return (x_center, item.Bottom);
            if (radio_bottom_right.Checked) return (item.Right, item.Bottom);

            return (item.Left, item.Right);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ProcessStartInfo psInfo = new ProcessStartInfo
            {
                FileName = "https://github.com/yevhenii-sir",
                UseShellExecute = true
            };
            Process.Start(psInfo);
        }

        private void uncheckAllBtn_Click(object sender, EventArgs e)
        {
            foreach (var node in Descendants(treeView1.Nodes))
                node.Checked = false;
        }

        private void checkAllBtn_Click(object sender, EventArgs e)
        {
            foreach (var node in Descendants(treeView1.Nodes))
                node.Checked = true;
        }

        private void openFileBtn_Click(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            openFileDialog1.Filter = "Photoshop Files|*.psd;|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path_file = openFileDialog1.FileName;

                try
                {
                    document = PsdDocument.Create(path_file);

                    label8.Text = "Document size: " + document.Width + " x " + document.Height;

                    List<string> temp_names_list = new List<string>();

                    treeView1.Nodes.Add(ShowLayers(document.Childs, "path", temp_names_list));
                    treeView1.ExpandAll();

                    textBox3.Text = path_file;
                }
                catch { }
            }
        }

        private void openFolderBtn_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    floder_to_save_image = fbd.SelectedPath;
                    textBox4.Text = floder_to_save_image;
                }
            }
        }

        private void saveImagesBtn_Click(object sender, EventArgs e)
        {
            if (floder_to_save_image != null)
            {
                try
                {
                    foreach (var node in Descendants(treeView1.Nodes))
                        if (node.Checked && node.Nodes.Count == 0)
                            SaveImageToFile(GetBitmap(PSDLayersDictionary[node.Text]), floder_to_save_image + "/" + node.Text, Convert.ToDouble(textBox9.Text));

                    MessageBox.Show("Saved!");
                    Process.Start("explorer.exe", floder_to_save_image);
                }
                catch { }
            }
            else MessageBox.Show("Сhoose a folder to save!");
        }

        private void getCoordinatesBtn_Click(object sender, EventArgs e)
        {
            try
            {
                textBox2.Text = "";
                double sceneScale = Convert.ToDouble(textBox8.Text);
                foreach (var node in Descendants(treeView1.Nodes))
                {
                    if (node.Checked && node.Nodes.Count == 0)
                    {
                        string tempString = "";

                        foreach (string item in listBox1.Items)
                        {
                            (double x, double y) coordinates = GetCoordinateWithOrigin(PSDLayersDictionary[node.Text]);
                            tempString += item switch
                            {
                                "$$$sp_name" => node.Text,
                                "$$$x" => (coordinates.x * sceneScale) - (Convert.ToInt32(textBox6.Text) * sceneScale),
                                "$$$y" => (coordinates.y * sceneScale) - (Convert.ToInt32(textBox7.Text) * sceneScale),
                                "$$$alpha" => $"{PSDLayersDictionary[node.Text].Opacity:0.###}",
                                _ => item
                            };
                        }

                        textBox2.Text += tempString + "\r\n";
                    }
                }
            }
            catch { }
        }

        private void addSeparatorBtn_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add(textBox5.Text);
            textBox5.Text = "";
            SaveList();
        }

        private void clearListBtn_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 1)
            {
                (listBox1.Items[listBox1.SelectedIndex - 1], listBox1.Items[listBox1.SelectedIndex]) = (listBox1.Items[listBox1.SelectedIndex], listBox1.Items[listBox1.SelectedIndex - 1]);
                listBox1.SetSelected(Math.Clamp(listBox1.SelectedIndex - 1, 0, listBox1.Items.Count - 1), true);
                SaveList();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex <= (listBox1.Items.Count - 2) && listBox1.SelectedIndex >= 0)
            {
                (listBox1.Items[listBox1.SelectedIndex + 1], listBox1.Items[listBox1.SelectedIndex]) = (listBox1.Items[listBox1.SelectedIndex], listBox1.Items[listBox1.SelectedIndex + 1]);
                listBox1.SetSelected(Math.Clamp(listBox1.SelectedIndex + 1, 0, listBox1.Items.Count - 1), true);
                SaveList();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                SaveList();
            }
            catch { }
        }
    }
}
