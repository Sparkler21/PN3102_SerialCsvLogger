using System;
using System.IO;
using System.IO.Ports;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;

namespace SerialCsvLogger
{
    public partial class Form1 : Form
    {
        private SerialPort _port;
        private StreamWriter _writer;
        private string _csvPath;

        // UI
        private ComboBox portBox;
        private ComboBox baudBox;
        private Button startBtn;
        private Button stopBtn;
        private Button saveAsBtn;
        private Button clearBtn;
        private Label statusLabel;
        private RichTextBox liveBox;

        // outbound controls
        private TextBox sendBox;
        private Button sendBtn;
        private ComboBox lineEndBox; // None, \n, \r, \r\n
        private CheckBox echoCheck;

        public Form1()
        {
            InitializeComponent();
            BuildUi();
            LoadPorts();
        }

        private void BuildUi()
        {
            Text = "Serial → CSV Logger (WindSpeed/WindDirection)";
            Width = 900; Height = 620;

            portBox = new ComboBox { Left = 12, Top = 12, Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
            baudBox = new ComboBox { Left = 160, Top = 12, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            baudBox.Items.AddRange(new object[] { "9600", "19200", "38400", "57600", "115200", "230400" });
            baudBox.SelectedItem = "115200";

            saveAsBtn = new Button { Left = 270, Top = 10, Width = 110, Height = 28, Text = "Save CSV As…" };
            startBtn = new Button { Left = 390, Top = 10, Width = 80, Height = 28, Text = "Start" };
            stopBtn = new Button { Left = 475, Top = 10, Width = 80, Height = 28, Text = "Stop", Enabled = false };
            clearBtn = new Button { Left = 560, Top = 10, Width = 80, Height = 28, Text = "Clear" };

            statusLabel = new Label { Left = 12, Top = 46, Width = 860, Height = 20, Text = "Idle." };

            // Send row
            sendBox = new TextBox { Left = 12, Top = 72, Width = 590, TabIndex = 0 };
            lineEndBox = new ComboBox { Left = 610, Top = 72, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            lineEndBox.Items.AddRange(new object[] { "None", "\\n", "\\r", "\\r\\n" });
            lineEndBox.SelectedIndex = 1; // default \n
            sendBtn = new Button { Left = 715, Top = 70, Width = 70, Height = 26, Text = "Send" };
            echoCheck = new CheckBox { Left = 790, Top = 74, Width = 80, Text = "Echo", Checked = true };

            liveBox = new RichTextBox
            {
                Left = 12,
                Top = 104,
                Width = 860,
                Height = 470,
                ReadOnly = true,
                DetectUrls = false,
                WordWrap = false,
                Font = new Font("Consolas", 9)
            };

            Controls.AddRange(new Control[] {
                portBox, baudBox, saveAsBtn, startBtn, stopBtn, clearBtn, statusLabel,
                sendBox, lineEndBox, sendBtn, echoCheck,
                liveBox
            });

            saveAsBtn.Click += SaveAsBtn_Click;
            startBtn.Click += StartBtn_Click;
            stopBtn.Click += StopBtn_Click;
            clearBtn.Click += (s, e) => liveBox.Clear();

            sendBtn.Click += (s, e) => SendCurrentText();
            sendBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift && !e.Control && !e.Alt)
                {
                    e.Handled = true; e.SuppressKeyPress = true;
                    SendCurrentText();
                }
            };

            FormClosing += (s, e) => Cleanup();
            SetSendUiEnabled(false);
        }

        private void LoadPorts()
        {
            portBox.Items.Clear();
            foreach (var name in SerialPort.GetPortNames())
                portBox.Items.Add(name);
            if (portBox.Items.Count > 0) portBox.SelectedIndex = 0;
        }

        private void SaveAsBtn_Click(object sender, EventArgs e)
        {
            using var sfd = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                FileName = $"wind_{DateTime.Now:yyyy-MM-dd}.csv"
            };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                _csvPath = sfd.FileName;
                statusLabel.Text = $"CSV path set → {_csvPath}";
            }
        }

        private void StartBtn_Click(object sender, EventArgs e)
        {
            if (portBox.SelectedItem == null)
            {
                MessageBox.Show("Select a COM port first.", "Serial → CSV Logger");
                return;
            }
            if (string.IsNullOrWhiteSpace(_csvPath))
            {
                var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                _csvPath = Path.Combine(docs, $"wind_{DateTime.Now:yyyy-MM-dd}.csv");
                statusLabel.Text = $"No CSV chosen — using {_csvPath}";
            }

            int baud = int.Parse((string)baudBox.SelectedItem);

            bool newFile = !File.Exists(_csvPath) || new FileInfo(_csvPath).Length == 0;
            _writer = new StreamWriter(_csvPath, append: true, Encoding.UTF8);
            if (newFile) _writer.WriteLine("timestamp_iso,wind_speed,wind_direction");
            _writer.Flush();

            _port = new SerialPort((string)portBox.SelectedItem, baud, Parity.None, 8, StopBits.One)
            {
                NewLine = "\n",          // device sends \r\n — \n works fine as the terminator
                Encoding = Encoding.UTF8,
                ReadTimeout = 1000,
                WriteTimeout = 1000
            };
            _port.DataReceived += Port_DataReceived;

            try
            {
                _port.Open();
                statusLabel.Text = $"Logging on {_port.PortName} @ {baud} → {_csvPath}";
                startBtn.Enabled = false;
                stopBtn.Enabled = true;
                portBox.Enabled = false;
                baudBox.Enabled = false;
                saveAsBtn.Enabled = false;
                SetSendUiEnabled(true);
                sendBox.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open {_port.PortName}: {ex.Message}", "Error");
                Cleanup();
            }
        }

        private void StopBtn_Click(object sender, EventArgs e)
        {
            statusLabel.Text = "Stopping…";
            Cleanup();
            statusLabel.Text = "Stopped.";
            startBtn.Enabled = true;
            stopBtn.Enabled = false;
            portBox.Enabled = true;
            baudBox.Enabled = true;
            saveAsBtn.Enabled = true;
            SetSendUiEnabled(false);
        }

        // Regex to match: ,<speed>,<direction> (optionally with whitespace)
        // e.g. ",3.45,270"  ", 5 ,  123.4"
        private static readonly Regex WindRegex =
            new Regex(@"^\s*,\s*([+-]?\d+(?:\.\d+)?)\s*,\s*([+-]?\d+(?:\.\d+)?)\s*$",
                      RegexOptions.Compiled);

        private void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string raw = _port.ReadLine();                 // includes trailing \n (and maybe \r)
                string line = raw.TrimEnd('\r', '\n', ' ');    // clean line
                string ts = DateTime.Now.ToString("s");

                if (TryParseWind(line, out double ws, out double wd))
                {
                    // Write parsed fields to CSV
                    lock (this)
                    {
                        _writer.WriteLine($"{ts},{ws.ToString(CultureInfo.InvariantCulture)},{wd.ToString(CultureInfo.InvariantCulture)}");
                        _writer.Flush();
                    }

                    // Live view
                    BeginInvoke(new Action(() =>
                    {
                        TrimLiveBox();
                        liveBox.AppendText($"{ts}  WS={ws}  WD={wd}\n");
                        liveBox.ScrollToCaret();
                    }));
                }
                else
                {
                    // Not in expected format — show raw line but don't log to CSV
                    BeginInvoke(new Action(() =>
                    {
                        TrimLiveBox();
                        liveBox.AppendText($"{ts}  [unparsed] {line}\n");
                        liveBox.ScrollToCaret();
                        statusLabel.Text = "Received line not matching ,<WindSpeed>,<WindDirection>";
                    }));
                }
            }
            catch (TimeoutException) { /* ignore */ }
            catch (IOException)
            {
                BeginInvoke(new Action(() => statusLabel.Text = "I/O error (device removed?)."));
            }
            catch (InvalidOperationException) { /* port closed while reading */ }
        }

        private bool TryParseWind(string line, out double windSpeed, out double windDirection)
        {
            var m = WindRegex.Match(line);
            if (m.Success)
            {
                if (double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out windSpeed) &&
                    double.TryParse(m.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out windDirection))
                {
                    return true;
                }
            }
            windSpeed = 0; windDirection = 0;
            return false;
        }

        private void TrimLiveBox()
        {
            const int maxChars = 500_000;
            if (liveBox.TextLength > maxChars)
                liveBox.Select(0, liveBox.TextLength - maxChars);
        }

        // --- sending logic (unchanged) ---
        private void SendCurrentText()
        {
            try
            {
                if (_port == null || !_port.IsOpen)
                {
                    System.Media.SystemSounds.Beep.Play();
                    statusLabel.Text = "Not connected — press Start first.";
                    return;
                }

                var text = sendBox.Text;
                if (string.IsNullOrEmpty(text)) return;

                string suffix = lineEndBox.SelectedItem?.ToString() ?? "None";
                string terminator = suffix switch
                {
                    "\\n" => "\n",
                    "\\r" => "\r",
                    "\\r\\n" => "\r\n",
                    _ => string.Empty
                };

                string payload = text + terminator;
                _port.Write(payload);

                if (echoCheck.Checked)
                {
                    var ts = DateTime.Now.ToString("s");
                    liveBox.AppendText($"{ts}  >> {text}{(suffix == "None" ? "" : $" ({suffix})")}\n");
                    liveBox.ScrollToCaret();
                }

                sendBox.Clear();
                sendBox.Focus();
            }
            catch (TimeoutException)
            {
                statusLabel.Text = "Write timed out.";
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"Send failed: {ex.Message}";
            }
        }

        private static string EscapeCsv(string s)
        {
            if (s.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0)
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private void Cleanup()
        {
            try
            {
                if (_port != null)
                {
                    _port.DataReceived -= Port_DataReceived;
                    if (_port.IsOpen) _port.Close();
                    _port.Dispose();
                    _port = null;
                }
            }
            catch { /* ignore */ }

            try { _writer?.Dispose(); } catch { }
            _writer = null;
        }

        private void SetSendUiEnabled(bool on)
        {
            sendBox.Enabled = on;
            sendBtn.Enabled = on;
            lineEndBox.Enabled = on;
            echoCheck.Enabled = on;
        }
    }
}
