using System.Diagnostics;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;

namespace LehrerStundenPlan2ExcelWinForm;

public partial class Form1 : Form
{
    private Button btnConvert;
    private RichTextBox rtbJson;
    private string currentFilePath;

    public Form1()
    {
        Text = "Stundenplan JSON → Excel";
        Width = 1000;
        Height = 700;
        InitializeComponents();
    }

    private void InitializeComponents()
    {
        rtbJson = new RichTextBox
        {
            Multiline = true,
            ScrollBars = RichTextBoxScrollBars.Both,
            Dock = DockStyle.Fill,
            Font = new Font("Consolas", 10),
            ReadOnly = false,
            WordWrap = false
        };

        btnConvert = new Button
        {
            Text = "In Excel umwandeln",
            Dock = DockStyle.Bottom,
            Height = 40
        };

        Controls.Add(rtbJson);
        Controls.Add(btnConvert);

        AllowDrop = true;
        rtbJson.AllowDrop = true;

        DragEnter += Form1_DragEnter;
        DragDrop += Form1_DragDrop;
        rtbJson.DragEnter += Form1_DragEnter;
        rtbJson.DragDrop += Form1_DragDrop;

        btnConvert.Click += BtnConvert_Click;
    }

    private void Form1_DragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
            e.Effect = DragDropEffects.Copy;
    }

    private void Form1_DragDrop(object? sender, DragEventArgs e)
    {
        var files = (string[])e.Data.GetData(DataFormats.FileDrop);
        if (files.Length > 0 && File.Exists(files[0]))
        {
            currentFilePath = files[0];
            _ = LoadFileAsync(files[0]);
        }
    }

    private async Task LoadFileAsync(string path)
    {
        rtbJson.Clear();
        const int linesPerChunk = 500;
        var sb = new StringBuilder();

        // Lesen im Hintergrund, Append per Invoke um UI nicht zu blockieren
        await Task.Run(async () =>
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096,
                FileOptions.SequentialScan);
            using var sr = new StreamReader(fs, Encoding.UTF8);

            int count = 0;
            while (!sr.EndOfStream)
            {
                var line = await sr.ReadLineAsync().ConfigureAwait(false);
                sb.AppendLine(line);
                if (++count >= linesPerChunk)
                {
                    var chunk = sb.ToString();
                    sb.Clear();
                    count = 0;
                    rtbJson.Invoke((Action)(() => rtbJson.AppendText(chunk)));
                }
            }

            if (sb.Length > 0)
            {
                var remainder = sb.ToString();
                rtbJson.Invoke((Action)(() => rtbJson.AppendText(remainder)));
            }
        });
    }

    private void BtnConvert_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(rtbJson.Text))
        {
            MessageBox.Show("Bitte JSON einfügen oder Datei ziehen!", "Fehler",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            using var doc = JsonDocument.Parse(rtbJson.Text);
            var root = doc.RootElement;

            if (!root.TryGetProperty("days", out var days))
            {
                MessageBox.Show("JSON enthält kein 'days'-Element.", "Fehler",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Stundenplan");

            var dayList = days.EnumerateArray().ToList();
            var totalSlots = root.GetProperty("slots").GetArrayLength();
            var startCol = 2;

            ws.Cell(2, 1).Value = "Lehrkraft";
            var col = startCol;

            foreach (var day in dayList)
            {
                var date = day.GetProperty("day").GetString() ?? "???";
                var dayStartCol = col;

                for (var slot = 1; slot <= totalSlots; slot++)
                {
                    ws.Cell(2, col++).Value = slot;
                }

                var dayEndCol = col - 1;
                ws.Range(1, dayStartCol, 1, dayEndCol).Merge();
                ws.Cell(1, dayStartCol).Value = date;
                ws.Cell(1, dayStartCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(1, dayStartCol).Style.Font.Bold = true;
            }

            var teacherNames = new SortedSet<string>();
            foreach (var day in dayList)
            {
                foreach (var resource in day.GetProperty("resources").EnumerateArray())
                {
                    var name = resource.GetProperty("resource").GetProperty("shortName").GetString() ?? "?";
                    teacherNames.Add(name);
                }
            }

            var row = 3;
            foreach (var teacher in teacherNames)
            {
                ws.Cell(row, 1).Value = teacher;
                col = startCol;

                foreach (var day in dayList)
                {
                    var resource = day.GetProperty("resources")
                        .EnumerateArray()
                        .FirstOrDefault(r => r.GetProperty("resource").GetProperty("shortName").GetString() == teacher);

                    if (resource.ValueKind == JsonValueKind.Undefined)
                    {
                        for (var i = 0; i < totalSlots; i++)
                        {
                            ws.Cell(row, col++).Value = 0;
                        }

                        continue;
                    }

                    foreach (var cell in resource.GetProperty("cells").EnumerateArray())
                    {
                        var belegt = cell.GetProperty("gridEntries").GetArrayLength() > 0;
                        ws.Cell(row, col).Value = belegt ? 1 : string.Empty;
                        if (belegt)
                            ws.Cell(row, col).Style.Fill.BackgroundColor = XLColor.Gray;
                        col++;
                    }
                }

                row++;
            }

            ws.Column(1).AdjustToContents();
            var used = ws.RangeUsed();
            used.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            used.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            var slotWidth = 2.5;
            var slotStartCol = startCol;
            var slotEndCol = startCol + dayList.Count * totalSlots - 1;

            for (var c = slotStartCol; c <= slotEndCol; c++)
            {
                ws.Column(c).Width = slotWidth;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "Excel-Datei|*.xlsx",
                FileName = "stundenplan2.xlsx"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                wb.SaveAs(sfd.FileName);
                MessageBox.Show("Excel gespeichert:\n" + sfd.FileName, "Fertig",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Fehler: " + ex.Message, "Fehler",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}