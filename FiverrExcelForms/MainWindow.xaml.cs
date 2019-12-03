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
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Winforms = System.Windows.Forms;
using System.IO;

namespace FiverrExcelForms
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        ObservableCollection<ClosedPt> closedpt_list = new ObservableCollection<ClosedPt>();
        ObservableCollection<RefPt> Referentpt_list = new ObservableCollection<RefPt>();
        double startz = 0;
        double starte = 0;
        double startn = 0;
        double sumDist = 0;
        double totalDiffFromStartPoint = 0;
        int closedptNo = 0;
        int refptNo = 0;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void closed_loop_btn_Click(object sender, RoutedEventArgs e)
        {
            closedpt_list.Clear();
            ClosedDG.ItemsSource = null;
            Referentpt_list.Clear();
            referentDG.ItemsSource = null;
            if (!double.TryParse(StartZ_txt.Text, out startz) || !double.TryParse(StartE_txt.Text, out starte) || !double.TryParse(StartN_txt.Text, out startn) || !int.TryParse(closed_loop_pts.Text, out closedptNo))
            {
                MessageBox.Show("Check your inputs \n Should have only numbers");
            }
            else
            {
                ClosedPt firstpt = new ClosedPt { StartZ = startz, E = starte, N = startn };
                closedpt_list.Add(firstpt);
                for (int i = 1; i < closedptNo; i++) {
                    closedpt_list.Add(new ClosedPt());
                }
                ClosedDG.ItemsSource = closedpt_list;
            }
            closed_loopCompute_btn.IsEnabled = true;

            Referent_btn.IsEnabled = false;
            ReferentCompute_btn.IsEnabled = false;
            saveAs_btn.IsEnabled = false;
        }

        private void saveAs_btn_Click(object sender, RoutedEventArgs e)
        {
            Random rd = new Random();
            string nwfilename = @"\SurveyData" + rd.Next(5000) + ".csv";
            string newfile = @"";
            Winforms.FolderBrowserDialog dialog = new Winforms.FolderBrowserDialog();
            Winforms.DialogResult r = dialog.ShowDialog();
            if (r == Winforms.DialogResult.OK)
            {
                newfile = dialog.SelectedPath + nwfilename;
            }

            try
            {
                FileStream f = new FileStream(newfile, FileMode.Create);
                try
                {
                    using (StreamWriter sw = new StreamWriter(f))
                    {
                        sw.WriteLine("Closed Loop Point Table");
                        sw.WriteLine("Stand Point,Target Point,Degree,Minute,Second,Distance,Diff Elev(Delta Z),East(X),North(Y),Elevation(Z),CW,CCW,World Angle,StartZ,E,N,Sum Distance,Differential to start pt(X),Differential to start pt(Y),Differential to start pt(Z),Total Differential From Start Pt");
                        #region start writing
                        for (int i = 0; i < closedpt_list.Count; i++)
                        {
                            if (i == 0)
                            {
                                sw.Write(closedpt_list[i].Standpt + ",");
                                sw.Write(closedpt_list[i].TargetPoint + ",");
                                sw.Write(closedpt_list[i].Degree + ",");
                                sw.Write(closedpt_list[i].Minute + ",");
                                sw.Write(closedpt_list[i].Second + ",");
                                sw.Write(closedpt_list[i].Distance + ",");
                                sw.Write(closedpt_list[i].DeltaZ + ",");
                                sw.Write(closedpt_list[i].X + ",");
                                sw.Write(closedpt_list[i].Y + ",");
                                sw.Write(closedpt_list[i].Z + ",");
                                sw.Write(closedpt_list[i].Cw + ",");
                                sw.Write(closedpt_list[i].Ccw + ",");
                                sw.Write(closedpt_list[i].WorldA + ",");
                                sw.Write(closedpt_list[i].StartZ + ",");
                                sw.Write(closedpt_list[i].E + ",");
                                sw.Write(closedpt_list[i].N + ",");
                                sw.Write(closedpt_list[i].Sum_distance + ",");
                                sw.Write(closedpt_list[i].Diff_x + ",");
                                sw.Write(closedpt_list[i].Diff_y + ",");
                                sw.Write(closedpt_list[i].Diff_z + ",");
                                sw.Write(totalDiffFromStartPoint.ToString());
                                sw.WriteLine();
                            }
                            else
                            {
                                sw.Write(closedpt_list[i].Standpt + ",");
                                sw.Write(closedpt_list[i].TargetPoint + ",");
                                sw.Write(closedpt_list[i].Degree + ",");
                                sw.Write(closedpt_list[i].Minute + ",");
                                sw.Write(closedpt_list[i].Second + ",");
                                sw.Write(closedpt_list[i].Distance + ",");
                                sw.Write(closedpt_list[i].DeltaZ + ",");
                                sw.Write(closedpt_list[i].X + ",");
                                sw.Write(closedpt_list[i].Y + ",");
                                sw.Write(closedpt_list[i].Z + ",");
                                sw.Write(closedpt_list[i].Cw + ",");
                                sw.Write(closedpt_list[i].Ccw + ",");
                                sw.Write(closedpt_list[i].WorldA + ",");
                                sw.Write(closedpt_list[i].StartZ + ",");
                                sw.Write(closedpt_list[i].E + ",");
                                sw.Write(closedpt_list[i].N + ",");
                                sw.Write(closedpt_list[i].Sum_distance + ",");
                                sw.Write(closedpt_list[i].Diff_x + ",");
                                sw.Write(closedpt_list[i].Diff_y + ",");
                                sw.Write(closedpt_list[i].Diff_z + ",");
                                sw.WriteLine();
                            }


                        }
                        #endregion

                        sw.WriteLine();
                        sw.WriteLine();
                        sw.WriteLine("Referent Point Table");
                        sw.WriteLine("Stand Point,Target Point,Degree,Minute,Second,Distance,Diff Elev(Delta Z),East(X),North(Y),Elevation(Z),CW,CCW,World Angle,E,N,Differential to start pt(Z)");


                        for (int i = 0; i < Referentpt_list.Count; i++)
                        {
                            sw.Write(Referentpt_list[i].RefPoint.TargetPoint + ",");
                            sw.Write(Referentpt_list[i].TargetPoint + ",");
                            sw.Write(Referentpt_list[i].Degree + ",");
                            sw.Write(Referentpt_list[i].Minute + ",");
                            sw.Write(Referentpt_list[i].Second + ",");
                            sw.Write(Referentpt_list[i].Distance + ",");
                            sw.Write(Referentpt_list[i].DeltaZ + ",");
                            sw.Write(Referentpt_list[i].X + ",");
                            sw.Write(Referentpt_list[i].Y + ",");
                            sw.Write(Referentpt_list[i].Z + ",");
                            sw.Write(Referentpt_list[i].Cw + ",");
                            sw.Write(Referentpt_list[i].Ccw + ",");
                            sw.Write(Referentpt_list[i].WorldA + ",");
                            sw.Write(Referentpt_list[i].E + ",");
                            sw.Write(Referentpt_list[i].N + ",");
                            sw.Write(Referentpt_list[i].Diff_z);
                            sw.WriteLine();
                        }
                    }
                    f.Close();
                    this.ShowMessageAsync("Saved", "Path location of saved file:" + newfile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            

           

        }

        private void Calc_ClosedLoop()
        {
            //processing for the first row 
            
            closedpt_list[0].Cw = closedpt_list[0].Degree + (closedpt_list[0].Minute / 60) + (closedpt_list[0].Second / 3600);
            closedpt_list[0].Ccw = 360 - closedpt_list[0].Cw;
            closedpt_list[0].WorldA = 0;
            closedpt_list[0].Sum_distance = 0;
            closedpt_list[0].Diff_x = 0;
            closedpt_list[0].Diff_y = 0;
            closedpt_list[0].Diff_z = 0;
            closedpt_list[0].X = closedpt_list[0].E;
            closedpt_list[0].Y = closedpt_list[0].N;
            closedpt_list[0].Z = closedpt_list[0].StartZ;

            //processing for the other rows
            for (int i = 1; i < closedpt_list.Count; i++)
            {
                if (i == 1)
                {
                    //closedpt_list[i].StandPoint = closedpt_list[i].StandPoint;
                    closedpt_list[i].Cw = closedpt_list[i].Degree + (closedpt_list[i].Minute / 60) + (closedpt_list[i].Second / 3600);
                    closedpt_list[i].Ccw = 360 - closedpt_list[i].Cw;
                    closedpt_list[i].WorldA = closedpt_list[0].Ccw;
                    closedpt_list[i].StartZ = closedpt_list[i - 1].StartZ + closedpt_list[i].DeltaZ;
                    closedpt_list[i].E=closedpt_list[i-1].E +(closedpt_list[i].Distance*(Math.Cos(closedpt_list[i].WorldA*(Math.PI/180))));
                    closedpt_list[i].N = closedpt_list[i - 1].N + (closedpt_list[i].Distance * (Math.Sin(closedpt_list[i].WorldA * (Math.PI / 180))));
                    closedpt_list[i].Sum_distance = closedpt_list[i - 1].Sum_distance + closedpt_list[i].Distance;
                }
                else
                {
                    //closedpt_list[i].StandPoint = closedpt_list[i].StandPoint;
                    closedpt_list[i].Cw = closedpt_list[i].Degree + (closedpt_list[i].Minute / 60) + (closedpt_list[i].Second / 3600);
                    closedpt_list[i].Ccw = 360 - closedpt_list[i].Cw;
                    closedpt_list[i].WorldA = (((closedpt_list[i - 1].Ccw + closedpt_list[i - 1].WorldA) - 180) % 360);
                    closedpt_list[i].StartZ = closedpt_list[i - 1].StartZ + closedpt_list[i].DeltaZ;
                    closedpt_list[i].E = closedpt_list[i - 1].E + (closedpt_list[i].Distance * (Math.Cos(closedpt_list[i].WorldA * (Math.PI / 180))));
                    closedpt_list[i].N = closedpt_list[i - 1].N + (closedpt_list[i].Distance * (Math.Sin(closedpt_list[i].WorldA * (Math.PI / 180))));
                    closedpt_list[i].Sum_distance = closedpt_list[i - 1].Sum_distance + closedpt_list[i].Distance;
                }

                
            }
            //processing for diffX,diffY,diffZ
            //for X
            closedpt_list[closedpt_list.Count - 1].Diff_x = closedpt_list[0].E - closedpt_list[closedpt_list.Count - 1].E;
            //for y
            closedpt_list[closedpt_list.Count - 1].Diff_y = closedpt_list[0].N - closedpt_list[closedpt_list.Count - 1].N;
            //for z
            closedpt_list[closedpt_list.Count - 1].Diff_z = closedpt_list[0].StartZ - closedpt_list[closedpt_list.Count - 1].StartZ;

            //processing the other outputs depending on the results of the just concluded calculations
            for (int j = 1; j < closedpt_list.Count-1; j++)
            {
                closedpt_list[j].Diff_x = (closedpt_list[j].Sum_distance / closedpt_list[closedpt_list.Count - 1].Sum_distance) * closedpt_list[closedpt_list.Count - 1].Diff_x;
                closedpt_list[j].Diff_y = (closedpt_list[j].Sum_distance / closedpt_list[closedpt_list.Count - 1].Sum_distance) * closedpt_list[closedpt_list.Count - 1].Diff_y;
                closedpt_list[j].Diff_z = (closedpt_list[j].Sum_distance / closedpt_list[closedpt_list.Count - 1].Sum_distance) * closedpt_list[closedpt_list.Count - 1].Diff_z;
            }

            //processing the final outputs East(X) North(Y) Z
            for (int k = 0; k < closedpt_list.Count; k++)
            {
                closedpt_list[k].X = closedpt_list[k].E + closedpt_list[k].Diff_x;
                closedpt_list[k].Y = closedpt_list[k].N + closedpt_list[k].Diff_y;
                closedpt_list[k].Z = closedpt_list[k].StartZ + closedpt_list[k].Diff_z;
            }

            //processint total diff from start point
            totalDiffFromStartPoint = Math.Sqrt((closedpt_list[closedpt_list.Count - 1].Diff_x * closedpt_list[closedpt_list.Count - 1].Diff_x) + (closedpt_list[closedpt_list.Count - 1].Diff_y * closedpt_list[closedpt_list.Count - 1].Diff_y));
            totaldiff_txt.Text = totalDiffFromStartPoint.ToString();
            this.ShowMessageAsync("Total Differential From Start Point", totalDiffFromStartPoint.ToString());

            referentDG.ItemsSource = null;
            closedpt_combo.ItemsSource = closedpt_list;
        }

        private void referent_btn_Click(object sender, RoutedEventArgs e)
        {
            Referentpt_list.Clear();
            referentDG.ItemsSource = null;

            if (!int.TryParse(Referent_pts.Text, out refptNo) )
            {
                MessageBox.Show("Check your inputs \n Should have only numbers");
            }
            else
            {
                for (int i = 0; i < refptNo; i++)
                {
                    Referentpt_list.Add(new RefPt());
                }
                referentDG.ItemsSource = Referentpt_list;
            }
            ReferentCompute_btn.IsEnabled = true;
        }

        private void Calc_Refent()
        {
            for (int i = 0; i < Referentpt_list.Count; i++)
            {
                Referentpt_list[i].Cw = Referentpt_list[i].Degree + (Referentpt_list[i].Minute / 60) + (Referentpt_list[i].Second / 3600);
                Referentpt_list[i].Ccw = 360 - Referentpt_list[i].Cw;
                Referentpt_list[i].WorldA = (((Referentpt_list[i].RefPoint.WorldA- Referentpt_list[i].Cw) + 180) % 360);
                Referentpt_list[i].E = Referentpt_list[i].RefPoint.E + Referentpt_list[i].RefPoint.Diff_x + (Referentpt_list[i].Distance * (Math.Cos(Referentpt_list[i].WorldA * (Math.PI / 180))));
                Referentpt_list[i].N = Referentpt_list[i].RefPoint.N + Referentpt_list[i].RefPoint.Diff_y + (Referentpt_list[i].Distance * (Math.Sin(Referentpt_list[i].WorldA * (Math.PI / 180))));
                Referentpt_list[i].Diff_z = Referentpt_list[i].RefPoint.StartZ + Referentpt_list[i].RefPoint.Diff_z + Referentpt_list[i].DeltaZ;
                Referentpt_list[i].X = Referentpt_list[i].E;
                Referentpt_list[i].Y = Referentpt_list[i].N;
                Referentpt_list[i].Z = Referentpt_list[i].Diff_z;

            }
        }
 
        private void closed_loopCompute_btn_Click(object sender, RoutedEventArgs e)
        {

            Calc_ClosedLoop();
            ClosedDG.ItemsSource = null;
            ClosedDG.ItemsSource = closedpt_list;
            Referent_btn.IsEnabled = true;
        }

        private void ReferentCompute_btn_Click(object sender, RoutedEventArgs e)
        {
            Calc_Refent();
            referentDG.ItemsSource = null;
            referentDG.ItemsSource = Referentpt_list;
            this.ShowMessageAsync("Calculation Complete", "Completed,you can save now");
            saveAs_btn.IsEnabled = true;
        }

       
    }
}
