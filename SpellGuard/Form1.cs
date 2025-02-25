using Microsoft.Office.Interop.Word;
using OfficeOpenXml;
using SpellGuard.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

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
                        errors = CheckSpell(selectedPath);
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

        private List<SpellError> CheckSpell(string folderPath)
        {
            List<SpellError> result = new List<SpellError>();

            if (!Directory.Exists(folderPath))
            {
                throw new Exception("Folder not found!");
            }

            var allFiles = Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories)
                                    .Where(s => s.EndsWith(".docx") || s.EndsWith(".doc"))
                                    .ToArray();
            progressBar1.Maximum = allFiles.Length;

            if (allFiles.Length == 0)
            {
                throw new Exception("No Word documents found in the folder.");
            }

            var wordApp = new Microsoft.Office.Interop.Word.Application();
            try
            {
                foreach (string filePath in allFiles)
                {
                    ProcessFile(filePath, wordApp, result);
                    progressBar1.Value += 1;
                }
            }
            finally
            {
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }

            return result;
        }

        private void ProcessFile(string filePath, Microsoft.Office.Interop.Word.Application wordApp, List<SpellError> result)
        {
            Document doc = null;
            try
            {
                doc = wordApp.Documents.Open(filePath);
                foreach (Microsoft.Office.Interop.Word.Range wordRange in doc.Words)
                {
                    if (wordRange.SpellingErrors.Count != 0)
                    {
                        var spellError = CreateSpellError(filePath, wordRange);
                        result.Add(spellError);
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                Console.WriteLine($"COM Error checking file {Path.GetFileName(filePath)}: {comEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking file {Path.GetFileName(filePath)}: {ex.Message}");
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(SaveChanges: false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }
            }
        }

        private SpellError CreateSpellError(string filePath, Microsoft.Office.Interop.Word.Range wordRange)
        {
            var pageNumber = wordRange.get_Information(WdInformation.wdActiveEndPageNumber);
            var lineNumber = wordRange.get_Information(WdInformation.wdFirstCharacterLineNumber);
            var position = wordRange.Words.First.Start;

            var suggestions = wordRange.GetSpellingSuggestions()
                                       .Cast<SpellingSuggestion>()
                                       .Select(s => s.Name)
                                       .ToList();

            return new SpellError
            {
                WordFileName = Path.GetFileName(filePath),
                WrongSpell = wordRange.Text,
                LineNumber = (int)lineNumber,
                PageNumber = (int)pageNumber,
                Position = position,
                SuggestedWords = String.Join(", ", suggestions)
            };
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
                    File.WriteAllBytes(saveFileDialog.FileName, content);
                    MessageBox.Show("File saved successfully.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private byte[] ExportExcel(List<SpellError> errors)
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

        private void DisplayErrors(List<SpellError> errors)
        {
            dataGridViewErrors.DataSource = errors;
        }

        private void Form1_Load(object sender, EventArgs e) { }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e) { }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e) { }
    }
}
