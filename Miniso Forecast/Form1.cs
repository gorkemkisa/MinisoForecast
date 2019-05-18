/******
 * 
 * İstanbul Bilgi Univercity 
 * 
 * Industrial Engineering - Senior Design Project 2019
 * 
 * Created by: Görkem Kısa
 * 
 * ****/


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RDotNet;
using System.Configuration;
using System.Threading;

namespace Miniso_Forecast
{
    public partial class Form1 : Form
    {

        string result;
        string DosyaYolu, DosyaAdi, SayfaAdi;

        int fore_aralik, fore_harici;
        string fore_cikis, fore_checkbox;

        int asso_harici, asso_sutun, asso_satir;
        string asso_cikis;

        int harici_int, harici_imp;
        string harici_cikis, harici_sayfaekle, harici_kod;

        REngine engine;


        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {

            Control.CheckForIllegalCrossThreadCalls = false;
            yukleme.Visible = false;

            // Checking if the R is installed on the computer

            try
            {

                //init the R engine  

                REngine.SetEnvironmentVariables();
                engine = REngine.GetInstance();
                engine.Initialize();
            }
            catch
            {
                MessageBox.Show("R bilgisayarınızda yüklü değil, R yükleyip programı tekrar çalıştırın","Hata");
                Application.Exit();
            }


        }

        private void f_calistir_Click(object sender, EventArgs e)
        {

            SayfaAdi = sayfaAdiBox.Text;

            // Checking whether the sheet name and the file path is correctly entered

            if (!check(DosyaYolu) && !check(SayfaAdi))
            {

                yukleme.Visible = true;

                //  Creating new thread for the R process

                new Thread(() =>
                {

                    // Gathered information is transferred onto the data

                    fore_cikis = f_cikissayfa.Text;
                    fore_aralik = Convert.ToInt32(f_aralik.Value);
                    fore_harici = Convert.ToInt32(f_haricietken.Value);

                    if (f_sayfaeklesinmi.Checked)
                    {
                        fore_checkbox = "addWorksheet(wb, sheetName = '" + fore_cikis + "');";
                    }
                    else
                    {
                        fore_checkbox = "";
                    }

                    DosyaYolu = DosyaYolu.Replace(@"\", "/");

                    string kod1 = "options(warn = 2);using <- function(...) {libs <- unlist(list(...));req <- unlist(lapply(libs,require,character.only=TRUE));need <- libs[req==FALSE];if(length(need)>0){ install.packages(need);lapply(need,require,character.only=TRUE);}};using('forecast', 'openxlsx');Sales <- read.xlsx('" + DosyaYolu + "',sheet='" + SayfaAdi + "');wb <- loadWorkbook('" + DosyaYolu + "');counter <- 1;rowindex <- 2;forecastperiod <- " + fore_aralik + "; m <- " + fore_harici + "/100;maxid <- length(Sales[,1]);" + fore_checkbox + "n <- Sales[,1];tSales <- as.data.frame(t(Sales[,-1]));colnames(tSales) <- n;tSales$myfactor <- factor(row.names(tSales));tSales$myfactor <- NULL;dummy <- data.frame(matrix(NA, nrow=maxid,ncol=forecastperiod+1));dummy[,1] <- Sales[,1];colnames(dummy) <- c('ID', seq(from = length(tSales[,1])+1, to = length(tSales[,1])+forecastperiod, by=1));writeData(wb, sheet='" + fore_cikis + "', dummy, startRow=1, startCol = 1);tsSales <- ts(tSales);forecastop <- tryCatch({while(counter <= maxid){forecast <- as.data.frame(thetaf(tsSales[,counter],level=95,h=forecastperiod));fc <- forecast[,1];fc[fc < 0] <- 0;fc <- fc+fc*m;forecastvalues <- round(fc,0);forecastvalues <- t(forecastvalues);writeData(wb, sheet='" + fore_cikis + "', forecastvalues,startRow = rowindex, startCol = 2, colNames = FALSE);counter <- counter+1;rowindex <- rowindex+1;};print('Forecasting successful. Now writing data to Excel file.')},error = function(e){print('Error. Problem in the R code. Check the error message:');stop(e);},warning = function(w){print('Warning. There may be a problem within the data. Check the error message:');stop(w);});writingop <- tryCatch({saveWorkbook(wb, '" + DosyaYolu + "', overwrite = TRUE);print('Done');},error = function(e){print('Error. Please check the error message:'); stop(e);},warning = function(w){stop('Access denied. Make sure that the Excel file is closed and not protected.');});openXL('" + DosyaYolu + "');print('Done')";

                    // Commencing the R operation

                    try
                    {
                        CharacterVector vector = engine.Evaluate(kod1).AsCharacter();
                        result = vector[0];

                        if (result == "Done")
                        {
                            MessageBox.Show("İşlem başarıyla tamamlandı");
                        }
                        else
                        {
                            MessageBox.Show(result);
                        }
                    }
                    catch (RDotNet.EvaluationException a)
                    {
                        MessageBox.Show(" " + a);

                    }

                    yukleme.Visible = false;
                    //clean up
                    //engine.Dispose();

                }).Start();

            }
            else
            {
                MessageBox.Show("Dosya yolu veya sayfa adı girilmedi !" + check(DosyaYolu) + " - " + check(SayfaAdi));
            }

        }

        private void a_calistir_Click(object sender, EventArgs e)
        {

            SayfaAdi = sayfaAdiBox.Text;


            if (!check(DosyaYolu) && !check(SayfaAdi))
            {

                yukleme.Visible = true;

                new Thread(() =>
                {

                    //calculate

                    asso_cikis = a_cikissayfasi.Text;
                    asso_sutun = Convert.ToInt32(a_sutun.Value);
                    asso_satir = Convert.ToInt32(a_satir.Value);
                    asso_harici = Convert.ToInt32(a_haricietken.Value);


                    DosyaYolu = DosyaYolu.Replace(@"\", "/");

                    string akod = "using <- function(...) {libs <- unlist(list(...));req <- unlist(lapply(libs,require,character.only=TRUE));need <- libs[req==FALSE];if(length(need)>0){;install.packages(need);lapply(need,require,character.only=TRUE);};};using('forecast', 'openxlsx');Sales <- read.xlsx('" + DosyaYolu + "',sheet='" + SayfaAdi + "');wb <- loadWorkbook('" + DosyaYolu + "');counter <- 1;rowindex <- " + asso_satir + "; colindex <- " + asso_sutun + "; m <- " + asso_harici + "/100; maxid <- length(Sales[,1]);n <- Sales[,1];tSales <- as.data.frame(t(Sales[,-1]));colnames(tSales) <- n;tSales$myfactor <- factor(row.names(tSales));tSales$myfactor <- NULL;tsSales <- ts(tSales);forecastop <- tryCatch({;while(counter <= maxid){;forecast <- as.data.frame(thetaf(tsSales[,counter],level=80,h=1));fc <- forecast[,1];fc[fc < 0] <- 0;fc <- fc+fc*m;forecastvalues <- round(fc,0);writeData(wb,sheet='" + asso_cikis + "',forecastvalues,startRow=rowindex,startCol = colindex);counter <- counter+1;rowindex <- rowindex+1;};print('Forecast values for assortment are generated. Now writing data to Excel file.');},error = function(e){print('Error. Problem in the R code. Check the error message:');stop(e);},warning = function(w){print('Warning. There may be a problem within the data. Check the warning message:');stop(w);});writingop <- tryCatch({saveWorkbook(wb, '" + DosyaYolu + "', overwrite = TRUE);print('Done');},error = function(e){;print('Error. Please check the error message:');stop(e);},warning = function(w){;stop('Access denied. Make sure that the Excel file is closed and not protected.');});openXL('" + DosyaYolu + "');print('Done')";


                    try
                    {
                        CharacterVector vector = engine.Evaluate(akod).AsCharacter();
                        result = vector[0];

                        if (result == "Done")
                        {
                            MessageBox.Show("İşlem başarıyla tamamlandı");
                        }
                        else
                        {
                            MessageBox.Show(result);
                        }
                    }
                    catch (RDotNet.EvaluationException a)
                    {
                        MessageBox.Show(" " + a);

                    }

                    yukleme.Visible = false;
                    //clean up
                    //engine.Dispose();

                }).Start();

            }
            else
            {
                MessageBox.Show("Dosya yolu veya sayfa adı girilmedi !" + check(DosyaYolu) + " - " + check(SayfaAdi));
            }

        }


        private void m_calistir_Click(object sender, EventArgs e)
        {

            SayfaAdi = sayfaAdiBox.Text;

            if (!check(DosyaYolu) && !check(SayfaAdi))
            {

                yukleme.Visible = true;

                DosyaYolu = DosyaYolu.Replace(@"\", "/");

                harici_cikis = m_cikissayfasi.Text;


                if (m_cikissayfacheck.Checked)
                {
                    harici_sayfaekle = "addWorksheet(wb, sheetName = '" + harici_cikis + "');";
                }
                else
                {
                    harici_sayfaekle = "";
                }


                if (m_remove.Checked)
                {

                    harici_kod = "using <- function(...) {libs <- unlist(list(...));req <- unlist(lapply(libs,require,character.only=TRUE));need <- libs[req==FALSE];if(length(need)>0){install.packages(need);lapply(need,require,character.only=TRUE);};};using('openxlsx', 'imputeTS');Sales <- read.xlsx('" + DosyaYolu + "',sheet='" + SayfaAdi + "');wb <- loadWorkbook('" + DosyaYolu + "');rowindex <- 2;maxid <- length(Sales[,1]);" + harici_sayfaekle + "n <- Sales[,1];tSales <- as.data.frame(t(Sales[,-1]));colnames(tSales) <- n;tSales$myfactor <- factor(row.names(tSales));tSales$myfactor <- NULL;for(i in 1:maxid){;ntSales <- na.remove(tSales[,i]);ntSales <- as.data.frame(t(ntSales));writeData(wb, sheet='" + harici_cikis + "', ntSales, startRow=rowindex, startCol = 2, colNames = FALSE);rowindex <- rowindex+1};writingop <- tryCatch({saveWorkbook(wb, '" + DosyaYolu + "', overwrite = TRUE);print('Done');},error = function(e){print('Error. Please check the error message:');stop(e);},warning = function(w){stop('Access denied. Make sure that the Excel file is closed and not protected.');});openXL('" + DosyaYolu + "');print('Done')";

                }
                else if (m_katman.Checked)
                {

                    harici_kod = "using <- function(...) {libs <- unlist(list(...));req <- unlist(lapply(libs,require,character.only=TRUE));need <- libs[req==FALSE];if(length(need)>0){ ;install.packages(need);lapply(need,require,character.only=TRUE);};};using('openxlsx', 'imputeTS');Sales <- read.xlsx('" + DosyaYolu + "',sheet='" + SayfaAdi + "');wb <- loadWorkbook('" + DosyaYolu + "');maxid <- length(Sales[,1]);" + harici_sayfaekle + "n <- Sales[,1];tSales <- as.data.frame(t(Sales[,-1]));colnames(tSales) <- n;tSales$myfactor <- factor(row.names(tSales));tSales$myfactor <- NULL;ntSales <- na.kalman(tSales);ntSales[ntSales < 0] <- 0;ntSales <- round(ntSales,0);ntSales <- as.data.frame(t(ntSales));writeData(wb, sheet='" + harici_cikis + "', ntSales, startRow=2, startCol = 2, colNames = FALSE);writingop <- tryCatch({saveWorkbook(wb, '" + DosyaYolu + "', overwrite = TRUE);print('Done');},error = function(e){print('Error. Please check the error message:');stop(e);},warning = function(w){stop('Access denied. Make sure that the Excel file is closed and not protected.');});openXL('" + DosyaYolu + "');print('Done')";

                }
                else if (m_mice.Checked)
                {
                    harici_int = Convert.ToInt32(m_int.Value);
                    harici_imp = Convert.ToInt32(m_imp.Value);
                    harici_kod = "using <- function(...) {;libs <- unlist(list(...));req <- unlist(lapply(libs,require,character.only=TRUE));need <- libs[req==FALSE];if(length(need)>0){ ;install.packages(need);lapply(need,require,character.only=TRUE);};};using('openxlsx', 'mice');Sales <- read.xlsx('" + DosyaYolu + "',sheet='" + SayfaAdi + "');wb <- loadWorkbook('" + DosyaYolu + "');rowindex <- 2;imptime <- " + harici_imp + ";inttime <- " + harici_int + ";maxid <- length(Sales[,1]);" + harici_sayfaekle + "n <- Sales[,1];tSales <- as.data.frame(t(Sales[,-1]));colnames(tSales) <- n;tSales$myfactor <- factor(row.names(tSales));tSales$myfactor <- NULL;colnames(tSales) <- c();tsSales <- ts(tSales);imputed_Data <- mice(tsSales, m=imptime, maxit = inttime, method = 'pmm');completeData <- complete(imputed_Data, 1);writeData(wb, sheet='" + harici_cikis + "', ntSales, startRow=rowindex, startCol = 2, colNames = FALSE);writingop <- tryCatch({;saveWorkbook(wb, '" + DosyaYolu + "', overwrite = TRUE);print('Done');},error = function(e){;print('Error. Please check the error message:');stop(e);},warning = function(w){;stop('Access denied. Make sure that the Excel file is closed and not protected.');});openXL('" + DosyaYolu + "');print('Done')";

                }


                new Thread(() =>
                {


                    if (m_mice.Checked)
                    {
                        DialogResult cikis = new DialogResult();
                        cikis = MessageBox.Show("Bu işlem girilen değerlere bağlı olarak uzun sürebilir, devam etmek istiyor musunuz ?", "Uyarı", MessageBoxButtons.YesNo);
                        if (cikis == DialogResult.Yes)
                        {

                            try
                            {
                                CharacterVector vector = engine.Evaluate(harici_kod).AsCharacter();
                                result = vector[0];

                                if (result == "Done")
                                {
                                    MessageBox.Show("İşlem başarıyla tamamlandı");
                                }
                                else
                                {
                                    MessageBox.Show(result);
                                }
                            }
                            catch (RDotNet.EvaluationException a)
                            {
                                MessageBox.Show(" " + a);

                            }
                        }
                        if (cikis == DialogResult.No)
                        {



                        }
                    }
                    else
                    {

                        try
                        {
                            CharacterVector vector = engine.Evaluate(harici_kod).AsCharacter();
                            result = vector[0];

                            if (result == "Done")
                            {
                                MessageBox.Show("İşlem başarıyla tamamlandı");
                            }
                            else
                            {
                                MessageBox.Show(result);
                            }
                        }
                        catch (RDotNet.EvaluationException a)
                        {
                            MessageBox.Show(" " + a);

                        }


                    }



                    yukleme.Visible = false;
                    //clean up
                    //engine.Dispose();

                }).Start();



            }
            else
            {
                MessageBox.Show("Dosya yolu veya sayfa adı girilmedi !" + check(DosyaYolu) + " - " + check(SayfaAdi));
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            MessageBox.Show("Hazırlayan: Görkem Kısa");


        }


        private void m_mice_CheckedChanged(object sender, EventArgs e)
        {

            if (m_mice.Checked)
            {

                m_int.Enabled = true;
                m_imp.Enabled = true;

            }
            else
            {
                m_int.Enabled = false;
                m_imp.Enabled = false;
            }

        }

        private void dosyasec_Click(object sender, EventArgs e)
        {

            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel Dosyası |*.xlsx;*.xlsm;*.xls";
            file.FilterIndex = 2;
            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyası Seçiniz..";
            file.Multiselect = true;

            if (file.ShowDialog() == DialogResult.OK)
            {
                DosyaYolu = file.FileName;
                DosyaAdi = file.SafeFileName;
                label2.Text = DosyaYolu;
            }

        }

        public static bool check(string s)
        {
            return (s == null || s == String.Empty) ? true : false;
        }

        private void yardim1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Zaman cinsi, veride yer alan zaman cinsiyle aynı olmalıdır.", "Bilgi");
        }

        private void yardim2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Yüzde cinsinden forecastler %+/- 200'e kadar arttırıp azaltılabilir. Satışlarda ani bir artma/azalma öngörüyorsanız bunu kullanabilirsiniz.", "Bilgi");
        }

        private void yardim3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Hangi Excel sayfasına çıktı alınacak? Eğer Excel sayfası yoksa, alttaki seçeneği tıklamayı unutmayın.", "Bilgi");
        }

        private void yardim4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Satır index ve sütun index excelde hangi satır ve sütundan başlayarak doldurulacağını seçebilirsiniz.", "Bilgi");
        }

        private void yardim5_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Yüzde cinsinden forecastler %+/- 200'e kadar arttırıp azaltılabilir. Satışlarda ani bir artma/azalma öngörüyorsanız bunu kullanabilirsiniz.", "Bilgi");
        }

        private void yardim6_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Bu modül yeni bir Excel sayfası eklemez, sadece varolan bir Excel sayfasına işlem yapar.", "Bilgi");
        }

        private void yardim11_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Verisetindeki veriler arasındaki boşlukları siler, ve veriyi kaydırır. Az veri varsa dikkatli kullanın.", "Bilgi");
        }

        private void yardim12_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Kalman Smoothing metotuyla boşlukları doldurur.", "Bilgi");
        }

        private void yardim13_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Mice algoritması kullanarak doldurur. Küçük verisetlerinde, küçük imp. time ve int. time ile yapılması tavsiye edilir. Minimum yarım saat sürer.", "Bilgi");
        }

        private void yardim15_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Imp. time ve int. time ne kadar fazla olursa, doğruluk ve harcanan zaman da o kadar fazla olur.", "Bilgi");
        }

        private void yardim21_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Bu modül ile Excel'de yer alan satış verisine belirlenen forecast aralığı kadar tahmin yapılabilir.", "Bilgi");
        }

        private void yardim22_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Bu modül forecast modülü ile aynı özelliklere sahiptir. Tek farkı, 1 zaman birimi bazında tahmin yapar.", "Bilgi");
        }

        private void yardim14_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Hangi Excel sayfasına çıktı alınacak? Eğer Excel sayfası yoksa, alttaki seçeneği tiklemeyi unutmayın.", "Bilgi");
        }


    }
}
