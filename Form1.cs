using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using QRCoder;

namespace PrintPromo
{
    public partial class Form1 : Form
    {
        private BindingSource bindingSource1 = new BindingSource();

        public Form1()
        {
            InitializeComponent();
        }

        private void selectedButon()
        {
            var selectedRows = dataGridView1.Rows
                .Cast<DataGridViewRow>()
                .Where(r => r.Cells["Selected"].Value != null && Convert.ToBoolean(r.Cells["Selected"].Value))
                .ToList();

            if (selectedRows.Count == 0)
            {
                toolStripButton3.Image = global::PrintPromo.Properties.Resources.checkmarksquare_120277;
                toolStripButton3.Text = "Выделить все";
                toolStripLabel1.Text = "0 записей выбрано";
            }
            else
            {
                toolStripButton3.Image = global::PrintPromo.Properties.Resources.minussquare_120267;
                toolStripButton3.Text = "Снять все выделеные";
                toolStripLabel1.Text = $"{selectedRows.Count} записей выбрано";
            }

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Tool strip button clicked!");
        }
        private void Filter_Click(object sender, EventArgs e)
        {
            // Пример фильтрации по названию акции
            bindingSource1.Filter = $"name LIKE '%{toolStripTextBox1.Text}%' or promoNo LIKE '%{toolStripTextBox1.Text}%'";
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback +=
                   (s, cert, chain, sslPolicyErrors) => true;
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            string dateFilter = DateTime.Today.ToString("yyyy-MM-dd");
            string url = $"https://cucrm.kz/api/v1/CPromoTest?where[0][type]=greaterThan&where[0][attribute]=endDate&where[0][value]={dateFilter}";

            string login = "admin";
            string password = "YaoulierKing210895";
            string credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{login}:{password}"));

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", credentials);

                var response = await client.GetStringAsync(url);
                var promoResponse = JsonConvert.DeserializeObject<PromoResponse>(response);

                // Преобразуем список в DataTable для фильтрации
                System.Data.DataTable table = ToDataTable(promoResponse.list);
                bindingSource1.DataSource = table;
                dataGridView1.DataSource = bindingSource1;

                if (dataGridView1.Columns["Selected"] == null)
                {
                    DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                    checkBoxColumn.Name = "Selected";
                    checkBoxColumn.HeaderText = "Выбрать";
                    checkBoxColumn.FalseValue = false;
                    checkBoxColumn.TrueValue = true;
                    checkBoxColumn.DataPropertyName = "";
                    dataGridView1.Columns.Insert(0, checkBoxColumn);
                }
            }
        }



        public static System.Data.DataTable ToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            System.Data.DataTable table = new System.Data.DataTable(); // Explicitly specify System.Data.DataTable
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

       

private void ExportSelectedToWord()
        {
            // Блокируем форму и показываем курсор выполнения  
            this.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                var selectedRows = dataGridView1.Rows
                    .Cast<DataGridViewRow>()
                    .Where(r => r.Cells["Selected"].Value != null && Convert.ToBoolean(r.Cells["Selected"].Value))
                    .ToList();

                if (selectedRows.Count == 0)
                {
                    MessageBox.Show("Не выбрано ни одной записи.");
                    return;
                }

                var wordApp = new Word.Application();
                wordApp.Visible = false;
                var doc = wordApp.Documents.Add();

                doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                doc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(2.7f);
                doc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1.0f);
                doc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.0f);
                doc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1.5f);

                int labelsPerPage = 16;
                float labelWidth = wordApp.CentimetersToPoints(6.5f);
                float labelHeight = wordApp.CentimetersToPoints(3.7f);

                int tableRows = 4, tableCols = 4;
                int count = 0;
                Word.Table table = null;

                for (int i = 0; i < selectedRows.Count; i++)
                {
                    if (count % labelsPerPage == 0)
                    {
                        if (table != null)
                        {
                            doc.Paragraphs.Add();
                        }

                        table = doc.Tables.Add(doc.Paragraphs.Last.Range, tableRows, tableCols);
                        table.Borders.Enable = 0;

                        for (int r = 1; r <= tableRows; r++)
                        {
                            for (int c = 1; c <= tableCols; c++)
                            {
                                table.Cell(r, c).Width = labelWidth;
                                table.Cell(r, c).Height = labelHeight;
                                table.Cell(r, c).HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                                table.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                            }
                        }

                        table.Spacing = 0;
                        table.AllowAutoFit = false;
                    }

                    int row = (count % labelsPerPage) / tableCols + 1;
                    int col = (count % labelsPerPage) % tableCols + 1;

                    var cell = table.Cell(row, col);
                    var data = selectedRows[i].Cells;

                    // Заполняем ячейку с форматированием  
                    FillLabelCell(cell, data);

                    count++;
                }

                wordApp.Visible = true;
                doc.Activate();
            }
            finally
            {
                // Разблокируем форму и возвращаем стандартный курсор  
                this.Enabled = true;
                Cursor.Current = Cursors.Default;
            }
        }

        private void FillLabelCell(Word.Cell cell, DataGridViewCellCollection data)
        {
            Word.Application wordApp = cell.Range.Application;

            var range = cell.Range;
            range.End -= 1; // исключаем знак конца ячейки

            string gdsNm = data["gdsNm"]?.Value?.ToString() ?? "";
            string gdsNm2 = data["gdsNm2"]?.Value?.ToString() ?? "";
            string percent = data["percentDiscount"]?.Value?.ToString() ?? "";
            string price = data["price"]?.Value?.ToString() ?? "";
            string priceDisc = data["priceDisc"]?.Value?.ToString() ?? "";
            string gdsCd = data["gdsCd"]?.Value?.ToString() ?? "";

            range.Text = "";

            // Выровняем всю ячейку по вертикали по центру
            cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            var lines = new[]
            {
        new { Text = gdsNm2, Size = 8, Bold = true,   Color =Word.WdColor.wdColorBlack, Align = Word.WdParagraphAlignment.wdAlignParagraphCenter, Strike = false },
        new { Text = gdsNm,  Size = 8, Bold = true,  Color = Word.WdColor.wdColorBlack, Align = Word.WdParagraphAlignment.wdAlignParagraphCenter, Strike = false },
        new { Text = $"-{percent}%", Size = 10, Bold = true, Color = Word.WdColor.wdColorRed, Align = Word.WdParagraphAlignment.wdAlignParagraphRight, Strike = false },
        new { Text =  $"{price}₸.",  Size = 9, Bold = true,  Color = Word.WdColor.wdColorBlue, Align = Word.WdParagraphAlignment.wdAlignParagraphRight, Strike = true } // горизонтальный центр
    };

            // ---- обычные строки ----
            foreach (var line in lines)
            {
                range.InsertAfter(line.Text);
                FormatLastInsertedText(range, cell, line);
                range.InsertParagraphAfter();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            // ---- последняя строка: QR + priceDisc ----
            Word.Table innerTable = cell.Range.Tables.Add(range, 1, 2);
            innerTable.Borders.Enable = 0;
            innerTable.Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            innerTable.Rows.Height = wordApp.CentimetersToPoints(2.5f);
            innerTable.Columns[1].Width = wordApp.CentimetersToPoints(2.5f); // ширина под QR
            innerTable.Columns[2].Width = wordApp.CentimetersToPoints(3.8f); // под цену
            innerTable.AllowAutoFit = false;

            Word.Cell qrCell = innerTable.Cell(1, 1);
            Word.Cell priceCell = innerTable.Cell(1, 2);

            // вертикальное и горизонтальное центрирование для ячеек таблицы
            qrCell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            priceCell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            priceCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // вставляем QR как изображение
            InsertQrImage(qrCell.Range, gdsCd);

            // вставляем цену
            Word.Range priceRange = priceCell.Range;
            priceRange.End -= 1;
            priceRange.Text =  $"{priceDisc}₸.";
            priceRange.Font.Size = 18;
            priceRange.Font.Bold = 1;
        }



        private void FormatLastInsertedText(Word.Range range, Word.Cell cell, dynamic line)
        {
            cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Word.Range textRange = cell.Range;
            textRange.Start = range.End - line.Text.Length ;
            textRange.End = range.End ;

            textRange.Font.Size = line.Size;
            textRange.Font.Bold = line.Bold ? 1 : 0;
            textRange.Font.Color = line.Color;
            textRange.Font.StrikeThrough = line.Strike ? 1 : 0;
            textRange.ParagraphFormat.Alignment = line.Align;
            textRange.ParagraphFormat.SpaceAfter = 0;
            textRange.ParagraphFormat.SpaceBefore = 0;
        }
        private void InsertQrImage(Word.Range targetRange, string text)
        {
            using (var qrGenerator = new QRCodeGenerator())
            {
                var qrData = qrGenerator.CreateQrCode(text, QRCodeGenerator.ECCLevel.Q);
                using (var qrCode = new QRCode(qrData))
                using (var bitmap = qrCode.GetGraphic(5)) // базовый размер
                {
                    string tempPath = System.IO.Path.GetTempFileName() + ".png";
                    bitmap.Save(tempPath);

                    var shape = targetRange.InlineShapes.AddPicture(tempPath, LinkToFile: false, SaveWithDocument: true);

                    // Уменьшаем размер QR на 30%
                    shape.Width = shape.Width * 0.7f;
                    shape.Height = shape.Height * 0.7f;

                    Word.Range afterQrRange = targetRange.Duplicate;
                    afterQrRange.Start = targetRange.End + 1;
                    afterQrRange.End = afterQrRange.Start;

                    // Добавляем перенос строки
                    afterQrRange.InsertParagraphAfter();
                    afterQrRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    // Добавляем подпись под QR
                    string caption = text?.Trim() ?? "";
                    if (caption.Length > 0)
                    {
                        afterQrRange.InsertAfter(caption);

                        // Форматируем подпись
                        Word.Range captionRange = afterQrRange.Duplicate;
                        captionRange.Start = afterQrRange.End - caption.Length;
                        captionRange.End = afterQrRange.End;

                        captionRange.Font.Size = 6;
                        captionRange.Font.Bold = 0;
                        captionRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        captionRange.ParagraphFormat.SpaceBefore = 0;
                        captionRange.ParagraphFormat.SpaceAfter = 0;
                    }

                    // Удаляем временный файл
                    System.IO.File.Delete(tempPath);
                }
            }
        }


        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ExportSelectedToWord();
            selectedButon();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedButon();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            var selectedRows = dataGridView1.Rows
                .Cast<DataGridViewRow>()
                .Where(r => r.Cells["Selected"].Value != null && Convert.ToBoolean(r.Cells["Selected"].Value))
                .ToList();

           

                if (selectedRows.Count == 0)
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["Selected"].Value = true;
                    }
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["Selected"].Value = false;
                    }
                }
            selectedButon();


        }
    }


    public class PromoResponse
    {
        public int total { get; set; }
        public List<PromoItem> list { get; set; }
    }
    public class PromoItem
    {
        public string promoNo { get; set; }
        public string name { get; set; }
        public string barcode { get; set; }
        public string begDate { get; set; }

        public string endDate { get; set; }
        public string gdsCd { get; set; }
        public string gdsNm { get; set; }
        public string gdsNm1 { get; set; }
        public string gdsNm2 { get; set; }
        public int price { get; set; }
        public int discount { get; set; }
        public int priceDisc { get; set; }
        public int percentDiscount { get; set; }

    }

}
