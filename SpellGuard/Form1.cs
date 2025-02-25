using Microsoft.Office.Interop.Word;
using OfficeOpenXml;
using SpellGuard.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;

namespace SpellGuard
{
    public partial class Form1 : Form
    {
        public List<SpellError> errors;

        public Form1()
        {
            InitializeComponent();
            button2.Hide();
            errors = new List<SpellError>();
        }

        public List<SpellError> CheckSpell(string folderPath)
        {
            List<SpellError> result = new List<SpellError>();

            // Check if the folder exists
            if (!Directory.Exists(folderPath))
            {
                throw new Exception("Folder not found!");
            }

            // Get all .doc and .docx files in the folder
            var allFiles = Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories)
                                     .Where(s => s.EndsWith(".docx") || s.EndsWith(".doc"))
                                     .ToArray();
            progressBar1.Maximum = allFiles.Length;

            var count = 0;

            if (allFiles.Length == 0)
            {
                throw new Exception("No Word documents found in the folder.");
            }

            // Initialize Word application
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            try
            {
                foreach (string filePath in allFiles)
                {
                    count += 1;
                    // Open the Word document
                    Document doc = wordApp.Documents.Open(filePath);
                    try
                    {
                        foreach (Microsoft.Office.Interop.Word.Range wordRange in doc.Words)
                        {
                            if (wordRange.SpellingErrors.Count != 0)
                            {
                                var pageNumber = wordRange.get_Information(WdInformation.wdActiveEndPageNumber);
                                var lineNumber = wordRange.get_Information(WdInformation.wdFirstCharacterLineNumber);
                                var position = wordRange.Words.First.Start;

                                // Get spelling suggestions
                                List<string> suggestions = new List<string>();
                                foreach (Microsoft.Office.Interop.Word.SpellingSuggestion suggestion in wordRange.GetSpellingSuggestions())
                                {
                                    suggestions.Add(suggestion.Name);
                                }

                                // Create SpellError object
                                SpellError spellError = new SpellError()
                                {
                                    WordFileName = Path.GetFileName(filePath),
                                    WrongSpell = wordRange.Text,
                                    LineNumber = (int)lineNumber,
                                    PageNumber = (int)pageNumber,
                                    Position = position,
                                    SuggestedWords = String.Join(", ", suggestions)
                                };
                                result.Add(spellError);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error checking file {Path.GetFileName(filePath)}: {ex.Message}");
                    }
                    finally
                    {
                        // Close the document
                        progressBar1.Value += 1;
                        doc.Close(SaveChanges: false);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return result;
            }
            finally
            {
                // Quit the Word application
                wordApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select a folder";
                folderBrowserDialog.ShowNewFolderButton = true;

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderBrowserDialog.SelectedPath;
                    button1.Enabled = false;
                    button2.Hide();
                    progressBar1.Visible = true;
                    progressBar1.Minimum = 0;
                    progressBar1.Value = 0;

                    try
                    {
                        errors = CheckSpell(selectedPath.Replace("\\", "\\\\"));
                        MessageBox.Show("Spell Check Completed");
                        DisplayErrors(errors);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error: {ex.Message}");
                    }
                    finally
                    {
                        button1.Enabled = true;
                        button2.Show();
                        progressBar1.Value = 0;
                    }
                }
            }
        }

        public byte[] ExportExcel(List<SpellError> errors)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Spell Errors");
                worksheet.Cells["A1"].LoadFromCollection(errors, true);

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                return stream.ToArray();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                var content = ExportExcel(errors);
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Save Spell Errors"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    System.IO.File.WriteAllBytes(saveFileDialog.FileName, content);
                    MessageBox.Show("File saved successfully.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void DisplayErrors(List<SpellError> errors)
        {
            dataGridViewErrors.DataSource = errors;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}