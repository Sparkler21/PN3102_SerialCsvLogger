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
        // Port A: logging + live view (+ optional send)
        private SerialPort _port;
        private StreamWriter _writer;
        private string _csvPath;

        // Port B: commands to second device
        private SerialPort _port2;

        // --- UI: Port A (existing) ---
        private ComboBox portBox;
        private ComboBox baudBox;
        private Button startBtn;
        private Button stopBtn;
        private Button saveAsBtn;
        private Button clearBtn;
        private Label statusLabel;
        private RichTextBox liveBox;

        private TextBox sendBox;
        private Button sendBtn;
        private ComboBox lineEndBox;   // None, \n, \r, \r\n
        private CheckBox echoCheck;

        // --- UI: Port B (NEW) ---
        private ComboBox port2Box;
        private ComboBox baud2Box;
        private Button connect2Btn;
        private Button disconnect2Btn;
        private NumericUpDown intBox;
        private ComboBox sendModeBox;  // ASCII decimal, Single byte (0–255)
        private ComboBox lineEnd2Box;  // None, \n, \r, \r\n (for ASCII mode)
        private Button sendIntBtn;
        private CheckBox echo2Check;

        public Form1()
        {
            InitializeComponent();
            BuildUi();
            LoadPorts();
        }

        private void BuildUi()
        {
            Text = "Serial → CSV Logger (WindSpeed/WindDirection + 2nd Port Commands)";
            Width = 930; Height = 700;
            FormClosing += (s, e) => Cleanup();

            // --- Top row (Port A controls) ---
            portBox = new ComboBox { Left = 12, Top = 12, Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
            baudBox = new ComboBox { Left = 160, Top = 12, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            baudBox.Items.AddRange(new object[] { "9600", "19200", "38400", "57600", "115200", "230400" });
            baudBox.SelectedItem = "115200";

            saveAsBtn = new Button { Left = 270, Top = 10, Width = 120, Height = 28, Text = "Save CSV As…" };
            startBtn = new Button { Left = 395, Top = 10, Width = 80, Height = 28, Text = "Start" };
            stopBtn = new Button { Left = 480, Top = 10, Width = 80, Height = 28, Text = "Stop", Enabled = false };
            clearBtn = new Button { Left = 565, Top = 10, Width = 80, Height = 28, Text = "Clear" };

            statusLabel = new Label { Left = 12, Top = 46, Width = 890, Height = 20, Text = "Idle." };

            // --- Send row for Port A (existing) ---
            sendBox = new TextBox { Left = 12, Top = 72, Width = 600, TabIndex = 0 };
            lineEndBox = new ComboBox { Left = 615, Top = 72, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            lineEndBox.Items.AddRange(new object[] { "None", "\\n", "\\r", "\\r\\n" });
            lineEndBox.SelectedIndex = 1; // \n default
            sendBtn = new Button { Left = 720, Top = 70, Width = 75, Height = 26, Text = "Send A" };
            echoCheck = new CheckBox { Left = 800, Top = 74, Width = 100, Text = "Echo A", Checked = true };

            // --- NEW: Port B controls row ---
            port2Box = new ComboBox { Left = 12, Top = 104, Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
            baud2Box = new ComboBox { Left = 160, Top = 104, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            baud2Box.Items.AddRange(new object[] { "9600", "19200", "38400", "57600", "115200", "230400" });
            baud2Box.SelectedItem = "115200";
            connect2Btn = new Button { Left = 270, Top = 102, Width = 90, Height = 28, Text = "Connect B" };
            disconnect2Btn = new Button { Left = 365, Top = 102, Width = 95, Height = 28, Text = "Disconnect", Enabled = false };

            intBox = new NumericUpDown { Left = 468, Top = 104, Width = 100, Minimum = 0, Maximum = 65535, DecimalPlaces = 0 };
            sendModeBox = new ComboBox { Left = 572, Top = 104, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            sendModeBox.Items.AddRange(new object[] { "ASCII decimal", "Single byte (0–255)" });
            sendModeBox.SelectedIndex = 0;

            lineEnd2Box = new ComboBox { Left = 726, Top = 104, Width = 90, DropDownStyle = ComboBoxStyle.DropDownList };
            lineEnd2Box.Items.AddRange(new object[] { "None", "\\n", "\\r", "\\r\\n" });
            lineEnd2Box.SelectedIndex = 0; // None default (common for raw integer commands)

            sendIntBtn = new Button { Left = 820, Top = 102, Width = 80, Height = 28, Text = "Send B", Enabled = false };
            echo2Check = new CheckBox { Left = 820, Top = 134, Width = 80, Text = "Echo B", Checked = true, Enabled = false };

            // --- Live view box ---
            liveBox = new RichTextBox
            {
                Left = 12,
                Top = 160,
                Width = 890,
                Height = 480,
                ReadOnly = true,
                DetectUrls = false,
                WordWrap = false,
                Font = new Font("Consolas", 9)
            };

            Controls.AddRange(new Control[] {
                // Port A row
                portBox, baudBox, saveAsBtn, startBtn, stopBtn, clearBtn, statusLabel,
                // Port A send row
                sendBox, lineEndBox, sendBtn, echoCheck,
                // Port B row
                port2Box, baud2Box, connect2Btn, disconnect2Btn, intBox, sendModeBox, lineEnd2Box, sendIntBtn, echo2Check,
                // Live
                liveBox
            });

            // Wire events
            saveAsBtn.Click += SaveAsBtn_Click;
            startBtn.Click += StartBtn_Click;
            stopBtn.Click += StopBtn_Click;
            clearBtn.Click += (s, e) => liveBox.Clear();

            sendBtn.Click += (s, e) => SendCurrentTextA();
            sendBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift && !e.Control && !e.Alt)
                {
                    e.Handled = true; e.SuppressKeyPress = true;
                    SendCurrentTextA();
                }
            };

            connect2Btn.Click += Connect2Btn_Click;
            disconnect2Btn.Click += Disconnect2Btn_Click;
            sendIntBtn.Click += (s, e) => SendIntegerOnB();
            intBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift && !e.Control && !e.Alt)
                {
                    e.Handled = true; e.SuppressKeyPress = true;
                    SendIntegerOnB();
                }
            };

            SetSendUiEnabled(false); // Port A send disabled until connected
            SetPortBUiEnabled(false); // Port B send disabled until connected
        }

        private void LoadPorts()
        {
            portBox.Items.Clear();
            port2Box.Items.Clear();
            foreach (var name in SerialPort.GetPortNames())
            {
                portBox.Items.Add(name);
                port2Box.Items.Add(name);
            }
            if (portBox.Items.Count > 0) portBox.SelectedIndex = 0;
            if (port2Box.Items.Count > 0) port2Box.SelectedIndex = Math.Min(1, port2Box.Items.Count - 1); // pick a different one if available
        }

        // ----- CSV path selection (Port A) -----
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

        // ----- Start/Stop for Port A -----
        private void StartBtn_Click(object sender, EventArgs e)
        {
            if (portBox.SelectedItem == null)
            {
                MessageBox.Show("Select a COM port for Port A.", "Serial → CSV Logger");
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
                NewLine = "\n", // device sends \r\n; \n is fine as terminator
                Encoding = Encoding.UTF8,
                ReadTimeout = 1000,
                WriteTimeout = 1000
            };
            _port.DataReceived += PortA_DataReceived;

            try
            {
                _port.Open();
                statusLabel.Text = $"A: Logging on {_port.PortName} @ {baud} → {_csvPath}";
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
                MessageBox.Show($"Could not open Port A ({_port.PortName}): {ex.Message}", "Error");
                CleanupA();
            }
        }

        private void StopBtn_Click(object sender, EventArgs e)
        {
            statusLabel.Text = "A: Stopping…";
            CleanupA();
            statusLabel.Text = "A: Stopped.";
            startBtn.Enabled = true;
            stopBtn.Enabled = false;
            portBox.Enabled = true;
            baudBox.Enabled = true;
            saveAsBtn.Enabled = true;
            SetSendUiEnabled(false);
        }

        // ----- DataReceived for Port A (wind CSV) -----
        private static readonly Regex WindRegex =
            new Regex(@"^\s*,\s*([+-]?\d+(?:\.\d+)?)\s*,\s*([+-]?\d+(?:\.\d+)?)\s*$",
                      RegexOptions.Compiled);

        private void PortA_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string raw = _port.ReadLine();
                string line = raw.TrimEnd('\r', '\n', ' ');
                string ts = DateTime.Now.ToString("s");

                if (TryParseWind(line, out double ws, out double wd))
                {
                    // CSV write
                    lock (this)
                    {
                        _writer.WriteLine($"{ts},{ws.ToString(CultureInfo.InvariantCulture)},{wd.ToString(CultureInfo.InvariantCulture)}");
                        _writer.Flush();
                    }

                    // Live
                    BeginInvoke(new Action(() =>
                    {
                        TrimLiveBox();
                        liveBox.AppendText($"{ts}  WS={ws}  WD={wd}\n");
                        liveBox.ScrollToCaret();
                    }));
                }
                else
                {
                    BeginInvoke(new Action(() =>
                    {
                        TrimLiveBox();
                        liveBox.AppendText($"{ts}  [A unparsed] {line}\n");
                        liveBox.ScrollToCaret();
                        statusLabel.Text = "A: Received line not matching ,<WindSpeed>,<WindDirection>";
                    }));
                }
            }
            catch (TimeoutException) { }
            catch (IOException)
            {
                BeginInvoke(new Action(() => statusLabel.Text = "A: I/O error (device removed?)."));
            }
            catch (InvalidOperationException) { /* Port closed while reading */ }
        }

        private bool TryParseWind(string line, out double windSpeed, out double windDirection)
        {
            var m = WindRegex.Match(line);
            if (m.Success &&
                double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out windSpeed) &&
                double.TryParse(m.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out windDirection))
            {
                return true;
            }
            windSpeed = 0; windDirection = 0;
            return false;
        }

        // ----- Sending on Port A (unchanged) -----
        private void SendCurrentTextA()
        {
            try
            {
                if (_port == null || !_port.IsOpen)
                {
                    System.Media.SystemSounds.Beep.Play();
                    statusLabel.Text = "A: Not connected — press Start first.";
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

                _port.Write(text + terminator);

                if (echoCheck.Checked)
                {
                    var ts = DateTime.Now.ToString("s");
                    liveBox.AppendText($"{ts}  >>A {text}{(suffix == "None" ? "" : $" ({suffix})")}\n");
                    liveBox.ScrollToCaret();
                }

                sendBox.Clear();
                sendBox.Focus();
            }
            catch (TimeoutException)
            {
                statusLabel.Text = "A: Write timed out.";
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"A: Send failed: {ex.Message}";
            }
        }

        // =======================
        //       PORT  B
        // =======================
        private void Connect2Btn_Click(object sender, EventArgs e)
        {
            if (port2Box.SelectedItem == null)
            {
                MessageBox.Show("Select a COM port for Port B.", "Serial → CSV Logger");
                return;
            }

            // Optional: warn if same port as A
            if (_port != null && _port.IsOpen && string.Equals((string)port2Box.SelectedItem, _port.PortName, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Port B is the same as Port A. Port B is intended for a second device.", "Heads up");
                return;
            }

            int baud = int.Parse((string)baud2Box.SelectedItem);
            _port2 = new SerialPort((string)port2Box.SelectedItem, baud, Parity.None, 8, StopBits.One)
            {
                NewLine = "\n",
                Encoding = Encoding.UTF8,
                ReadTimeout = 1000,
                WriteTimeout = 1000
            };

            try
            {
                _port2.Open();
                statusLabel.Text = $"B: Connected on {_port2.PortName} @ {baud}";
                connect2Btn.Enabled = false;
                disconnect2Btn.Enabled = true;
                port2Box.Enabled = false;
                baud2Box.Enabled = false;
                SetPortBUiEnabled(true);
                intBox.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open Port B ({_port2.PortName}): {ex.Message}", "Error");
                CleanupB();
            }
        }

        private void Disconnect2Btn_Click(object sender, EventArgs e)
        {
            CleanupB();
            statusLabel.Text = "B: Disconnected.";
        }

        private void SendIntegerOnB()
        {
            try
            {
                if (_port2 == null || !_port2.IsOpen)
                {
                    System.Media.SystemSounds.Beep.Play();
                    statusLabel.Text = "B: Not connected — press Connect B first.";
                    return;
                }

                int value = (int)intBox.Value;

                if (sendModeBox.SelectedIndex == 0) // ASCII decimal
                {
                    string suffix = lineEnd2Box.SelectedItem?.ToString() ?? "None";
                    string terminator = suffix switch
                    {
                        "\\n" => "\n",
                        "\\r" => "\r",
                        "\\r\\n" => "\r\n",
                        _ => string.Empty
                    };

                    string payload = value.ToString(CultureInfo.InvariantCulture) + terminator;
                    _port2.Write(payload);

                    if (echo2Check.Checked)
                    {
                        var ts = DateTime.Now.ToString("s");
                        string suffixNote = (suffix == "None") ? "" : $" ({suffix})";
                        liveBox.AppendText($"{ts}  >>B (ASCII) {value}{suffixNote}\n");
                        liveBox.ScrollToCaret();
                    }
                }
                else // Single byte (0–255)
                {
                    if (value < 0 || value > 255)
                    {
                        MessageBox.Show("Single byte mode requires a value from 0 to 255.", "Range");
                        return;
                    }
                    byte b = (byte)value;
                    _port2.Write(new byte[] { b }, 0, 1);

                    if (echo2Check.Checked)
                    {
                        var ts = DateTime.Now.ToString("s");
                        liveBox.AppendText($"{ts}  >>B (byte) 0x{b:X2} ({value})\n");
                        liveBox.ScrollToCaret();
                    }
                }
            }
            catch (TimeoutException)
            {
                statusLabel.Text = "B: Write timed out.";
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"B: Send failed: {ex.Message}";
            }
        }

        // ----- Helpers -----
        private void TrimLiveBox()
        {
            const int maxChars = 500_000;
            if (liveBox.TextLength > maxChars)
                liveBox.Select(0, liveBox.TextLength - maxChars);
        }

        private static string EscapeCsv(string s)
        {
            if (s.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0)
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private void Cleanup()
        {
            CleanupA();
            CleanupB();
        }

        private void CleanupA()
        {
            try
            {
                if (_port != null)
                {
                    _port.DataReceived -= PortA_DataReceived;
                    if (_port.IsOpen) _port.Close();
                    _port.Dispose();
                    _port = null;
                }
            }
            catch { }
            try { _writer?.Dispose(); } catch { }
            _writer = null;
        }

        private void CleanupB()
        {
            try
            {
                if (_port2 != null)
                {
                    if (_port2.IsOpen) _port2.Close();
                    _port2.Dispose();
                    _port2 = null;
                }
            }
            catch { }
            connect2Btn.Enabled = true;
            disconnect2Btn.Enabled = false;
            port2Box.Enabled = true;
            baud2Box.Enabled = true;
            SetPortBUiEnabled(false);
        }

        private void SetSendUiEnabled(bool on)
        {
            sendBox.Enabled = on;
            sendBtn.Enabled = on;
            lineEndBox.Enabled = on;
            echoCheck.Enabled = on;
        }

        private void SetPortBUiEnabled(bool on)
        {
            intBox.Enabled = on;
            sendIntBtn.Enabled = on;
            sendModeBox.Enabled = on;
            lineEnd2Box.Enabled = on;
            echo2Check.Enabled = on;
        }
    }
}
