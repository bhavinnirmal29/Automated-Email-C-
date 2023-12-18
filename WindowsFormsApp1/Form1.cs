using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using WindowsFormsApp1.Properties;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private List<string> selectedFiles = new List<string>();
        private string rootFolderPath = "";
        private string selectedExcelFilePath = "";
        Employee e = new Employee();
        List<Employee> emp = new List<Employee>();

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //Root Folder Path Function
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Browse Files Function to display the files in listBox. The function has no use because the pdf files are selected from the root folder through coding.
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Multiselect = true;
                openFileDialog.Filter = "All Files|*.*";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //selectedFiles.Clear(); // Clear existing selections
                    selectedFiles.AddRange(openFileDialog.FileNames);
                    listBoxFiles.Items.Clear();
                    listBoxFiles.Items.AddRange(selectedFiles.ToArray());
                }
            }
        }

        bool ShouldAttach(Employee employee, string fileName)
        {
            return fileName.Contains(employee.AccountNumber);
        }
        List<Employee> empList = new List<Employee>();
        // Function to read data from an Excel file and return a List<Employee>
        private List<Employee> ReadDataFromExcelFile(string filePath)
        {
            

            FileInfo excelFile = new FileInfo(filePath);
            // Set the EPPlus LicenseContext
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // For non-commercial use

            using (ExcelPackage package = new ExcelPackage(excelFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // Start from row 2 (assuming header is in row 1)
                {
                    empList.Add(new Employee
                    {
                        Name = worksheet.Cells[row, 1].Value?.ToString(),
                        Email = worksheet.Cells[row, 3].Value?.ToString(),
                        AccountNumber = worksheet.Cells[row, 2].Value?.ToString()
                    });
                }
            }

            return empList;
        }
        private void send_Email_onClick_Click(object sender, EventArgs e)
        {
            string directoryPath = rootFolderPath; // Replace with actual directory path
            string subject = textBox3.Text;
            string rootfolderpath_1 = textBox1.Text;
            string body_text1 = textBox2.Text;
            if (string.IsNullOrEmpty(subject))
            {
                //Subject Text is Empty
                MessageBox.Show("Subject is Empty. Please enter valid subject", "Error");
            }
            else if (string.IsNullOrEmpty(rootfolderpath_1))
            {
                //Body Text is Empty
                MessageBox.Show("Select Valid Root Folder", "Error");
            }
            else if (string.IsNullOrEmpty(body_text1))
            {
                //Body Text is Empty
                MessageBox.Show("Body Text is empty", "Error");
            }
            else if (listBoxFiles.Items.Count == 0)
            {
                // ListBox is empty
                MessageBox.Show("The ListBox is empty.");
            }
            else
            {

            }

            List<Employee> emp = new List<Employee>
            {
                new Employee { Name = "Laura Alipio", Email = "lauraa@caryacalgary.ca", AccountNumber = "7363" },
                new Employee { Name = "Adam House", Email = "adamh@caryacalgary.ca", AccountNumber = "3545" },
                new Employee { Name = "Jason Lukman", Email = "jasonL@caryacalgary.ca", AccountNumber = "6742" }

            };

            List<string> temp = new List<string>();

            foreach (Employee em in empList)
            {
                temp.Add(em.AccountNumber);
            }


            Console.WriteLine("This is emp : " + emp);
            string body_text = "";
            if (Directory.Exists(directoryPath))
            {
                Console.WriteLine("This is emp : "+emp);
                string[] files = Directory.GetFiles(directoryPath);

                foreach (string filePath in files)
                {
                    string fileName = Path.GetFileName(filePath);

                   /* foreach (Employee employee in empList)
                    {*/
                        
                        if (fileName.Contains(employee.AccountNumber))
                        {
                            Console.WriteLine(employee.AccountNumber);
                            // Create an instance of the Outlook application
                            Outlook.Application outlookApp = new Outlook.Application();
                            // Create a new email item
                            Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                            // Set email properties
                            mailItem.Subject = textBox3.Text+" - "+employee.Name;
                            body_text = textBox2.Text;
                            DateTime today = DateTime.Today;
                            string body = @"
                            Hi,
                            <br><br>
                            Please submit the following by 12:00 noon on " + today.ToShortDateString() + @":
                            <br><br>
                            <b>a) Excel file of your BMO Reconciliation Summary</b><br>
                            <b>b) Copy of the attached Statement</b><br>
                            <b>c) PDF File of all the receipts ( Please combine all the receipts in just one file when scanning )</b>
                                <br>
                            <br>Below is the link for your reference to submit online the above mentioned documents to your Supervisor:
                            <br><br>
                            [Submission Link](https://forms.office.com/Pages/ResponsePage.aspx?id=Mna2vfbXakmTmbYFzrjkldikmSjo_2VFooJDayds46VUOUpUNTE5U1c5MjAwWDNWSkM2UE9DMFgyTy4u)
                            <br><br>
                            Thank you.<br>
                             <br<br>
                            " + @""+body_text;

                            // Signature HTML template
                            string signature = @"
                            <p>Regards,<br>
                            Laura Alipio<br>
                            Accountant<br>
                            <a href='ap@caryacalgary.ca'>ap@caryacalgary.ca</a><br>
                            Phone: 403-205-5231 </p>";

                            // Combine the main body and signature
                            string emailBody = $"{body}<br><br>{signature}";
                            
                            // Set the email body
                            mailItem.HTMLBody = emailBody;

                            mailItem.To = employee.Email;
                            // Attach the file
                            string attachmentFilePath = rootFolderPath + "\\"+fileName; // Replace with actual file path
                            mailItem.Attachments.Add(attachmentFilePath, Outlook.OlAttachmentType.olByValue, 1, ""+fileName);


                            // Send the email
                            mailItem.Send();
               
                            Console.WriteLine($"Email sent to {employee.Name} (Account: {employee.AccountNumber}) with attachment.");
                            // Release resources
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
                            outlookApp = null;

                            // This line helps to clean up the Outlook process
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
           
                    /*}*/
                }
            }
            else
            {
                MessageBox.Show("Directory does not exist.","Directory Error");
            }

        }

        private void label2_Click(object sender, EventArgs e)
        {
            //Subject Label
        }
        //Upload Buttong function to upload the files so that it displays in the text box and gives a message about files. 
        private void upload_onClick_Click(object sender, EventArgs e)
        {
            
            // Use the 'selectedFiles' list to access the selected file paths
            if (selectedFiles.Count > 0)
            {
                string fileList = string.Join("\n", selectedFiles);
                MessageBox.Show($"Uploaded files:\n{fileList}", "Files Uploaded");
            }
            else
            {
                MessageBox.Show("No files selected.", "Files Uploaded");
            }

        }

        private void set_Path_Click(object sender, EventArgs e)
        {
            while (true)
            {
                using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                {
                    if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedPath = folderBrowserDialog.SelectedPath;

                        if (IsValidRootFolder(selectedPath))
                        {
                            rootFolderPath = selectedPath;
                            textBox1.Text = $"Root Folder: {rootFolderPath}";
                            break; // Exit the loop when a valid root folder is selected
                        }
                        else
                        {
                            MessageBox.Show("Invalid root folder selected. Please choose a valid root folder.");
                        }
                    }
                    else
                    {
                        // The user canceled the operation, so you can handle it accordingly.
                        break; // Exit the loop if the user cancels
                    }
                }
            }
            List<string> directoryNames = new List<string>();

            try
            {
                if (Directory.Exists(rootFolderPath))
                {
                    string[] directories = Directory.GetDirectories(rootFolderPath);
                    foreach (string directory in directories)
                    {
                        directoryNames.Add(Path.GetFileName(directory));
                        string d_name = Path.GetFileName(directory);
                        Console.WriteLine(d_name);
                    }
                }
                else
                {
                    Console.WriteLine("Folder does not exist.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }


        }

        // Validation function to check if a folder is a valid root folder
        private bool IsValidRootFolder(string folderPath)
        {
            // Add your validation criteria here.
            // For example, you can check if the folder exists and contains certain files or subfolders.
            // Return true if it meets your criteria, otherwise return false.

            if (!Directory.Exists(folderPath))
            {
                return false; // The folder doesn't exist
            }

            // Add additional validation logic as needed

            return true;
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            //Set Root Folder Path Text Box
        }

        private void listBoxFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            //List Box Function
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //Body Text Box
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            //Subject Text Box
        }

        private void body_Label_Click(object sender, EventArgs e)
        {
            //Body Label
        }

        private void label1_Click_1(object sender, EventArgs e)
        {
            //Heading Label Function
        }

        //Excel FIle Button for Selecting Users
        private void excelFileButton_Click(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count == 0 && dataGridView1.Columns.Count == 0)
{
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Multiselect = false; // Set to false to allow selecting only one file
                    openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        selectedExcelFilePath = openFileDialog.FileName;
                        // You can now work with the selected Excel file (selectedFilePath)
                    }
                }
            }
            else
            {
                // DataGridView is not empty, so handle the logic accordingly.
                MessageBox.Show("File is Already Selected. Please clear to select again");
            }
            // Usage
            string filePath = selectedExcelFilePath; // Replace with your Excel file path
            emp = ReadDataFromExcelFile(filePath);
            // Configure the DataGridView columns
            dataGridView1.ColumnCount = 3;
            dataGridView1.Columns[0].Name = "Name";
            dataGridView1.Columns[1].Name = "Email";
            dataGridView1.Columns[2].Name = "AccountNumber";

            // Populate the DataGridView with user data
            foreach (var user in emp)
            {
                dataGridView1.Rows.Add(user.Name, user.Email, user.AccountNumber);
            }

        }



        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure to exit the programme ?", "Exit", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                Application.Exit();
            }
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            // Clear the DataGridView
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            
        }



        // Now, emp contains the data read from the Excel file
    }
}
