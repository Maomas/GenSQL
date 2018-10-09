/* Merci d'utiliser GenSQL.
 Ce programme permet de créer le code SQL d'une table de données créée sur un fichier Excel. */



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
//using System.Windows.Shapes;
using System.IO;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Application_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            textBox_title_coordinates.Text = "E3".Trim();
            textBox_table_title.Text = "TZS_BIJDR_VRI_AANVR".Trim();
        }


        private void button_writing_Click(object sender, RoutedEventArgs e)
        {
            string file_path = textBox_file.Text;


            if (textBox_title_coordinates.Text == string.Empty)
            {
                MessageBox.Show("Choisis d'abord la case du titre.", "Attention", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            else if (textBox_excel.Text == string.Empty)
            {
                MessageBox.Show("Choisis d'abord un fichier Excel à lire.", "Attention", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            else if (textBox_file.Text == string.Empty)
            {
                MessageBox.Show("Choisis d'abord un fichier texte à créer.", "Attention", MessageBoxButton.OK, MessageBoxImage.Warning);
            }



            else
            {
                string excel_path = textBox_excel.Text;
                string titleValue = textBox_table_title.Text;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excel_path);
                Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];

                try
                {
                    xlWorksheet = xlWorkbook.Sheets[titleValue];
                }

                catch
                {
                    MessageBox.Show("La page Excel " + titleValue + " n'existe pas !");
                }

                Excel.Range xlRange = xlWorksheet.UsedRange;

                File.WriteAllText(file_path, string.Empty);



                string contain = "";
                string titleCell = textBox_title_coordinates.Text;


                if (xlWorksheet.Range[titleCell].Value2 == null || Convert.ToString(xlWorksheet.Range[titleCell].Value2) != titleValue)
                {
                    MessageBox.Show("Aucun titre détecté !", "Attention !", MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBox_excel.Text = string.Empty;
                }

                else
                {

                    titleValue = Convert.ToString(xlWorksheet.Range[titleCell].Value2);


                    contain = "CREATE TABLE " + titleValue + " (\n";

                    int i = xlWorksheet.Range[titleCell].Row;
                    int j = xlWorksheet.Range[titleCell].Column;

                    string columnName = "";
                    string dataType = "";
                    string nullable = "";
                    string comments = "";
                    string codeType = "";

                    int cpt = 0;

                    for (cpt = 0; cpt <= xlWorksheet.UsedRange.Rows.Count; cpt++)
                    {

                        if (xlWorksheet.Cells[i + 3 + cpt, j + 1].Value2 != null && xlWorksheet.Cells[i + 3 + cpt, j + 1].Font.Strikethrough == true)
                        {
                            continue;
                        }

                        else if (xlWorksheet.Cells[i + 3 + cpt, j + 1].Value2 != null && xlWorksheet.Cells[i + 3 + cpt, j + 1].Font.Strikethrough == false)
                        {
                            columnName = Convert.ToString(xlWorksheet.Cells[i + 3 + cpt, j + 1].Value2);

                            dataType = Convert.ToString(xlWorksheet.Cells[i + 3 + cpt, j + 2].Value2).Replace(" ",string.Empty);

                            if (Convert.ToString(xlWorksheet.Cells[i + 3 + cpt, j + 4].Value2) == "Yes" || Convert.ToString(xlWorksheet.Cells[i + 2 + cpt, j + 4].Value2) == "YES") nullable = "";

                            else if (Convert.ToString(xlWorksheet.Cells[i + 3 + cpt, j + 4].Value2) == "No" || Convert.ToString(xlWorksheet.Cells[i + 2 + cpt, j + 4].Value2) == "NO") nullable = "NOT NULL";

                            comments = comments + "COMMENT ON COLUMN " + titleValue + "." + columnName + " IS " + "'" + (Convert.ToString(xlWorksheet.Cells[i + 3 + cpt, j + 3].Value2)) + "'" + "/" + "\n";

                            if (xlWorksheet.Cells[i + 3 + cpt, j + 1].Value2 == null && nullable=="") contain = contain + columnName + " " + dataType + "" + nullable + "\n";

                            else if (xlWorksheet.Cells[i + 3 + cpt, j + 1].Value2 != null && nullable == "") contain = contain + columnName + " " + dataType + "" + nullable + ",\n";

                            else if (xlWorksheet.Cells[i + 3 + cpt, j + 1].Value2 != null && nullable == "NOT NULL") contain = contain + columnName + " " + dataType + " " + nullable + ",\n";

                            else contain = contain + columnName + " " + dataType + " " + nullable + "\n";

                        }

                        else break;

                    }

                    File.AppendAllText(file_path, contain);

                    int constr_cpt = 0;

                    string[] constr_tab = new string[xlRange.Rows.Count];

                    try
                    {

                        for (cpt = 1; cpt <= xlRange.Rows.Count; cpt++)
                        {
                            if (xlWorksheet.Cells[cpt, j + 1].Value2 != null && Convert.ToString(xlWorksheet.Cells[cpt, j + 1].Value2).Trim() == "Constraints")
                            {

                                while (xlWorksheet.Cells[cpt + constr_cpt + 1, j + 1].Value2 != null)
                                {

                                    if (xlWorksheet.Cells[cpt + constr_cpt + 1, j + 1].Value2 != null && xlWorksheet.Cells[cpt + constr_cpt + 1, j + 1].Font.Strikethrough == true)
                                    {
                                        constr_cpt++;
                                    }


                                    else if (xlWorksheet.Cells[cpt + constr_cpt + 1, j + 1].Value2 != null && xlWorksheet.Cells[cpt + constr_cpt + 1, j + 1].Font.Strikethrough == false)
                                    {
                                        constr_tab[constr_cpt] = Convert.ToString(xlWorksheet.Cells[cpt + constr_cpt + 1, j + 1].Value2);

                                        if (xlWorksheet.Cells[cpt + constr_cpt + 2, j + 1].Value2 != null && constr_tab[constr_cpt].Contains("PK"))
                                        {
                                            File.AppendAllText(file_path, "CONSTRAINT " + constr_tab[constr_cpt] + " " + Convert.ToString(xlWorksheet.Cells[cpt + constr_cpt + 1, j + 2].Value2).Replace("on", string.Empty).ToUpper() + " USING INDEX TABLESPACE PEN_INDEX ,\n");
                                        }


                                        else if (xlWorksheet.Cells[cpt + constr_cpt + 2, j + 1].Value2 != null && constr_tab[constr_cpt].Contains("FK"))
                                        {
                                            File.AppendAllText(file_path, "CONSTRAINT " + constr_tab[constr_cpt] + " " + Convert.ToString(xlWorksheet.Cells[cpt + constr_cpt + 1, j + 2].Value2).Replace("on", string.Empty).Replace("referencing", "references").ToUpper() + " ,\n");
                                        }
                                        else if (xlWorksheet.Cells[cpt + constr_cpt + 2, j + 1].Value2 == null && constr_tab[constr_cpt].Contains("FK"))
                                        {
                                            File.AppendAllText(file_path, "CONSTRAINT " + constr_tab[constr_cpt] + " " + Convert.ToString(xlWorksheet.Cells[cpt + constr_cpt + 1, j + 2].Value2).Replace("on", string.Empty).Replace("referencing", "references").ToUpper() + "\n");
                                        }
                                        else if(xlWorksheet.Cells[cpt + constr_cpt + 1, j + 1].Value2 != null && constr_tab[constr_cpt].Contains("UK"))
                                        {
                                            File.AppendAllText(file_path, "CONSTRAINT " + constr_tab[constr_cpt] + " " + Convert.ToString(xlWorksheet.Cells[cpt + constr_cpt + 1, j + 2].Value2).Replace("Unique constraint on","unique").ToUpper() + "\n");
                                        }
                                        else if (xlWorksheet.Cells[cpt + constr_cpt + 1, j + 1].Value2 == null && constr_tab[constr_cpt].Contains("UK"))
                                        {
                                            File.AppendAllText(file_path, "CONSTRAINT " + constr_tab[constr_cpt] + " " + Convert.ToString(xlWorksheet.Cells[cpt + constr_cpt + 1, j + 2].Value2).Replace("Unique constraint on", "unique").ToUpper() + ",\n");
                                        }
                                        else
                                        {
                                            File.AppendAllText(file_path, "");
                                        }

                                        constr_cpt++;

                                    }
                                }


                            }
                        }

                    }

                    catch
                    {
                        MessageBox.Show("Veuillez compléter les cases manquantes du tableau Excel.", "Attention", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }


                    File.AppendAllText(file_path, ") TABLESPACE PEN_DATA\n\\");

                    string[] ind_tab = new string[xlRange.Rows.Count];
                    string ind_file_name, ind_file_path;
                    string ind_file_dir = Path.GetDirectoryName(file_path);


                    int ind_cpt = 1;

                    for (cpt = 1; cpt <= xlRange.Rows.Count; cpt++)
                    {

                        if (xlWorksheet.Cells[cpt, j + 1].Value2 != null && Convert.ToString(xlWorksheet.Cells[cpt, j + 1].Value2).Trim() == "Indexes")
                        {

                            while (xlWorksheet.Cells[cpt + ind_cpt + 1, j + 1].Value2 != null)
                            {
                                if (xlWorksheet.Cells[cpt + ind_cpt + 1, j + 1].Value2 != null && xlWorksheet.Cells[cpt + ind_cpt + 1, j + 1].Font.Strikethrough == true) ind_cpt++;
                                else if (xlWorksheet.Cells[cpt + ind_cpt + 1, j + 1].Value2 != null && xlWorksheet.Cells[cpt + ind_cpt + 1, j + 1].Font.Strikethrough == false)
                                {
                                    ind_tab[ind_cpt] = Convert.ToString(xlWorksheet.Cells[cpt + ind_cpt + 1, j + 1].Value2);

                                    ind_file_name = "create_index_" + ind_tab[ind_cpt] + ".sql";

                                    ind_file_path = Path.Combine(ind_file_dir, ind_file_name);

                                    File.WriteAllText(ind_file_path, "");

                                    File.WriteAllText(ind_file_path, "CREATE INDEX " + ind_tab[ind_cpt] + " ON " + titleValue + " " + Convert.ToString(xlWorksheet.Cells[cpt + ind_cpt + 1, j + 2].Value2) + "TABLESPACE PEN_INDEX /");

                                    ind_cpt++;
                                }



                            }
                        }
                    }



                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Marshal.ReleaseComObject(xlRange);
                    Marshal.ReleaseComObject(xlWorksheet);


                    string comments_file_name = "comment_table_" + titleValue + ".sql";
                    string sequence_file_name = "create_sequence_" + titleValue + ".sql";

                    string comments_file_dir = Path.GetDirectoryName(file_path);
                    string sequence_file_dir = Path.GetDirectoryName(file_path);

                    string comments_file_path = Path.Combine(comments_file_dir, comments_file_name);
                    string sequence_file_path = Path.Combine(sequence_file_dir, sequence_file_name);

                    File.WriteAllText(comments_file_path, "");
                    File.WriteAllText(sequence_file_path, "");

                    File.WriteAllText(comments_file_path, comments);
                    File.WriteAllText(sequence_file_path, "CREATE SEQUENCE " + titleValue + " /");

                    MessageBox.Show("Le code a bien été écrit dans le fichier " + file_path + ".", "Ecriture terminée", MessageBoxButton.OK, MessageBoxImage.Information);

                    xlWorkbook.Close(0);
                    Marshal.ReleaseComObject(xlWorkbook);

                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);

                }



            }

        }

        private void button_file_Click(object sender, RoutedEventArgs e)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "create_table_" + textBox_table_title.Text + ".sql";

            if (saveFileDialog.ShowDialog() == true)
            {
                string file_path = saveFileDialog.FileName;
                textBox_file.Text = saveFileDialog.FileName;
                File.WriteAllText(file_path, "");



            }






        }

        private void closing_button_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void button_excel_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog.ShowDialog() == true)
            {

                textBox_excel.Text = openFileDialog.FileName;
            }
        }
    }
}
