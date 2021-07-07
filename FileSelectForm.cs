using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;


namespace WordFormatter
{
	class FileSelectForm : Form
	{
		private TextBox textbox;
		private Button browse;
		//private TextBox textboxtemplate; //ver 3.0
		private Button format;
		public static String[] fileToOpen;
		public static String templateFileName;
		public FileSelectForm()
		{

			this.FormBorderStyle = FormBorderStyle.Fixed3D;
			Text = "Testcase Document Formatter";
			
						
			Label label = new Label
			{
				MaximumSize = new Size(200, 0),
				AutoSize = true,
				Text = "Select the input .docx file with the testcases to be formatted:",
				Location = new Point(15, 40)

			};


			textbox = new TextBox
			{
				Size = new Size(300, 100), //Width and Height
				Location = new Point(210, 40),
				AutoSize = true,				
				BorderStyle = BorderStyle.FixedSingle,
				WordWrap = true/*, //ver 3.0
				Multiline = true,
				//MaximumSize = new Size(200, 50),
				ScrollBars = ScrollBars.Both*/				
			};


			

			browse = new Button
			{
				Text = "Browse",
				TextAlign = ContentAlignment.MiddleCenter,
				Size = new Size(80, 24), //Width and Height
				Location = new Point(520, 40)				
			};

			browse.Click += new EventHandler(browse_Click);


			//ver 3.0
			/*Label labelForTemplate = new Label
			{
				MaximumSize = new Size(200, 0),
				AutoSize = true,
				Text = "Select the Test Protocol template file:",
				Location = new Point(15, 160)

			};

			textboxtemplate = new TextBox
			{
				Size = new Size(280, 100), //Width and Height
				Location = new Point(210, 160),
				AutoSize = true,
				BorderStyle = BorderStyle.FixedSingle,
				WordWrap = true
				//Multiline = true,
				//MaximumSize = new Size(200, 50),
				//ScrollBars = ScrollBars.Both
			};

			textboxtemplate.TextChanged += new EventHandler(enable_FormatButton); // ver 3.0

			Button browseTemplate = new Button
			{
				Text = "Browse",
				TextAlign = ContentAlignment.MiddleCenter,
				Size = new Size(80, 24), //Width and Height
				Location = new Point(520, 160)
			};

			browseTemplate.Click += new EventHandler(browseTemplate_Click);
*/
			format = new Button
			{
				Text = "Format",
				TextAlign = ContentAlignment.MiddleCenter,
				Size = new Size(80, 24),
				Location = new Point(210, 200),
				Enabled = false // ver 2.0

			};
			format.Click += new EventHandler(format_Click);



			ClientSize = new Size(620, 250); 
			this.Controls.Add(label);
			this.Controls.Add(textbox);
			this.Controls.Add(browse);
			//this.Controls.Add(labelForTemplate);
			//this.Controls.Add(textboxtemplate);
			//this.Controls.Add(browseTemplate);
			this.Controls.Add(format);
			


		}

		private void browse_Click(object sender, EventArgs e)
		{
			var FD = new OpenFileDialog();
			FD.Multiselect = true;
			if (FD.ShowDialog() == DialogResult.OK)
			{
				fileToOpen = FD.FileNames;
				int numberOfFiles = fileToOpen.Length;
				textbox.Clear();

				for (int i=0; i<numberOfFiles; i++)
				{

					textbox.Text += fileToOpen[i];
					if (i != numberOfFiles - 1)
					{
						textbox.Text += "," + Environment.NewLine;
					}
				}

				bool isDocx = true;
				foreach (String file in fileToOpen)
				{
					isDocx = isDocx && Path.GetExtension(file).Equals(".docx");
				}
				if (isDocx == true)
				{
					format.Enabled = true;
				}						
				else
				{
					format.Enabled = false;
					MessageBox.Show("One or many of the input files have an unsupported file format.\nPlease use only input files with \".docx\" extension and try again.");
				}
			}
		}

		//ver 3.0
		/*private void enable_FormatButton(object sender, EventArgs e)
		{
			format.Enabled = true;
			//textbox.MinimumSize = new Size(200, textbox.TextLength *100);
		}*/

		//ver 3.0
		/*private void browseTemplate_Click(object sender, EventArgs e)
		{
			var FD = new OpenFileDialog();
			FD.Multiselect = false;
			if (FD.ShowDialog() == DialogResult.OK)
			{
				templateFileName = FD.FileNames[0];
				textboxtemplate.Text += templateFileName;			

			}

			if(Path.GetExtension(file).Equals(".docx"))
			{
			format.Enabled = true;
			}
			else
			{
			format.Enabled = false;
			MessageBox.Show("The template file should have \".docx\" extension. Please correct the extension and try again.");
			}
		}*/

		private void format_Click(object sender, EventArgs e)
		{
			this.Dispose();
		}


	}
}
