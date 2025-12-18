using System;
using System.IO;
using System.IO.Ports;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;


namespace SerialCsvLogger
{
    public partial class Form1 : Form
    {
        // Port A: logging + live view (+ optional send)
        private SerialPort _port;
        private StreamWriter _writer;
        private string _csvPath;
        // Latest known motor angle (NaN = unknown)
        private double _motorAngleDeg = double.NaN;
        // Matches: "... now at 135.000°", "... now at -45 deg", "... now at 270 degrees (wrapped)"
        private static readonly Regex NowAtRegex =
            new Regex(@"now\s*at\s*([+-]?\d+(?:\.\d+)?)\s*(?:°|deg|degrees)?\b",
                      RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Port B: commands + now receive view
        private SerialPort _port2;

        // --- UI: Port A (logging) ---
        private ComboBox portBox;
        private ComboBox baudBox;
        private Button startBtn;
        private Button stopBtn;
        private Button saveAsBtn;
        private Button clearABtn;   // Clear live A
        private Label statusLabel;
        private RichTextBox liveBoxA;

        private TextBox sendBox;
        private Button sendBtn;
        private ComboBox lineEndBox;   // None, \n, \r, \r\n
        private CheckBox echoCheck;

        // --- Charts (Port A) ---
        private Chart speedChart;
        private Chart dirChart;
        private const int MaxChartPoints = 600; // keep ~10 minutes at 1 Hz

        // --- UI: Port B (commands + receive) ---
        private ComboBox port2Box;
        private ComboBox baud2Box;
        private Button connect2Btn;
        private Button disconnect2Btn;
        private NumericUpDown intBox;
        private ComboBox sendModeBox;  // ASCII decimal, Single byte (0–255)
        private ComboBox lineEnd2Box;  // None, \n, \r, \r\n (for ASCII mode)
        private Button sendIntBtn;
        private CheckBox echo2Check;
        private ComboBox prefixBox;    // "A" or "R"
        private Label prefixLabel;
        private Button zeroBtn;        // sends "Z"
        // Rolling buffer for Port B to handle chunked serial input (no newline required)
        private readonly StringBuilder _port2Buf = new StringBuilder(256);
        private const int Port2BufMax = 2048;
        private readonly object _port2BufLock = new object();

        private Button clearBBtn;      // Clear live B
        private RichTextBox liveBoxB;  // NEW: receive view for Port B


        public Form1()
        {
            InitializeComponent();
            BuildUi();
            LoadPorts();
        }

        private void BuildUi()
        {
            Text = "Serial → CSV Logger (Wind A + Command/Receive B + Charts)";
            Width = 1120;
            int desired = 1060;
            int maxH = Screen.FromControl(this).WorkingArea.Height - 60; // leave a little margin for taskbar
            Height = Math.Min(desired, Math.Max(700, maxH));             // never smaller than 700
            StartPosition = FormStartPosition.CenterScreen;
            FormClosing += (s, e) => Cleanup();

            // --- Top row (Port A controls) ---
            portBox = new ComboBox { Left = 12, Top = 12, Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
            baudBox = new ComboBox { Left = 160, Top = 12, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            baudBox.Items.AddRange(new object[] { "9600", "19200", "38400", "57600", "115200", "230400" });
            baudBox.SelectedItem = "115200";

            saveAsBtn = new Button { Left = 270, Top = 10, Width = 120, Height = 28, Text = "Save CSV As…" };
            startBtn = new Button { Left = 395, Top = 10, Width = 80, Height = 28, Text = "Start" };
            stopBtn = new Button { Left = 480, Top = 10, Width = 80, Height = 28, Text = "Stop", Enabled = false };
            clearABtn = new Button { Left = 565, Top = 10, Width = 90, Height = 28, Text = "Clear A" };

            statusLabel = new Label { Left = 12, Top = 46, Width = 1080, Height = 20, Text = "Idle." };

            // --- Send row for Port A ---
            sendBox = new TextBox { Left = 12, Top = 72, Width = 620, TabIndex = 0 };
            lineEndBox = new ComboBox { Left = 637, Top = 72, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            lineEndBox.Items.AddRange(new object[] { "None", "\\n", "\\r", "\\r\\n" });
            lineEndBox.SelectedIndex = 1; // \n default
            sendBtn = new Button { Left = 742, Top = 70, Width = 85, Height = 26, Text = "Send A" };
            echoCheck = new CheckBox { Left = 834, Top = 74, Width = 100, Text = "Echo A", Checked = true };

            // --- Port A live view label ---
            var aLabel = new Label { Left = 12, Top = 104, Width = 300, Height = 18, Text = "Port A — Live (parsed wind):" };

            // --- Port A live view (reduced height to make room for charts) ---
            liveBoxA = new RichTextBox
            {
                Left = 12,
                Top = 124,
                Width = 1080,
                Height = 160,
                ReadOnly = true,
                DetectUrls = false,
                WordWrap = false,
                Font = new Font("Consolas", 9)
            };

            // --- NEW: Charts for Port A ---
            speedChart = CreateChart("Wind Speed", "Time", "WindSpd");
            speedChart.Left = 12; speedChart.Top = 124 + 160 + 8; speedChart.Width = 1080; speedChart.Height = 180;
            speedChart.Series[0].Name = "WindSpeed";

            dirChart = CreateChart("Wind Direction", "Time", "WindDir");
            dirChart.Left = 12; dirChart.Top = speedChart.Top + speedChart.Height + 8; dirChart.Width = 1080; dirChart.Height = 180;
            dirChart.Series[0].Name = "WindDirection";
            var areaD = dirChart.ChartAreas[0];
            areaD.AxisY.Minimum = 0;
            areaD.AxisY.Maximum = 360;
            areaD.AxisY.Interval = 90;

            // --- Port B controls row (pushed down to make room for charts) ---
            int chartsBlockHeight = (120 + 8) + (120 + 8); // speed chart + spacing + dir chart + spacing
            int shift = chartsBlockHeight;                 // ~256px shift from your previous layout

            var bHeader = new Label { Left = 12, Top = 392 + shift, Width = 300, Height = 18, Text = "Port B — Command + Receive:" };

            port2Box = new ComboBox { Left = 12, Top = 414 + shift, Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
            baud2Box = new ComboBox { Left = 160, Top = 414 + shift, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            baud2Box.Items.AddRange(new object[] { "9600", "19200", "38400", "57600", "115200", "230400" });
            baud2Box.SelectedItem = "115200";
            connect2Btn = new Button { Left = 270, Top = 412 + shift, Width = 90, Height = 28, Text = "Connect B" };
            disconnect2Btn = new Button { Left = 365, Top = 412 + shift, Width = 95, Height = 28, Text = "Disconnect", Enabled = false };

            intBox = new NumericUpDown { Left = 468, Top = 414 + shift, Width = 100, Minimum = -360, Maximum = 360, DecimalPlaces = 0 };

            // Prefix ("A"/"R")
            prefixLabel = new Label { Left = 572, Top = 418 + shift, Width = 50, Height = 18, Text = "Prefix" };
            prefixBox = new ComboBox { Left = 620, Top = 414 + shift, Width = 50, DropDownStyle = ComboBoxStyle.DropDownList };
            prefixBox.Items.AddRange(new object[] { "A", "R" });
            prefixBox.SelectedIndex = 0;

            sendModeBox = new ComboBox { Left = 680, Top = 414 + shift, Width = 170, DropDownStyle = ComboBoxStyle.DropDownList };
            sendModeBox.Items.AddRange(new object[] { "ASCII decimal", "Single byte (0–255)" });
            sendModeBox.SelectedIndex = 0;

            lineEnd2Box = new ComboBox { Left = 854, Top = 414 + shift, Width = 90, DropDownStyle = ComboBoxStyle.DropDownList };
            lineEnd2Box.Items.AddRange(new object[] { "None", "\\n", "\\r", "\\r\\n" });
            lineEnd2Box.SelectedIndex = 3;

            sendIntBtn = new Button { Left = 948, Top = 412 + shift, Width = 70, Height = 28, Text = "Send B", Enabled = false };
            echo2Check = new CheckBox { Left = 1022, Top = 416 + shift, Width = 80, Text = "Echo B", Checked = true, Enabled = false };

            // Clear B / Zero + label
            clearBBtn = new Button { Left = 1015, Top = 446 + shift, Width = 87, Height = 24, Text = "Clear B", Enabled = false };
            zeroBtn = new Button { Left = 912, Top = 446 + shift, Width = 95, Height = 24, Text = "Zero", Enabled = false };
            var bRecvLabel = new Label { Left = 12, Top = 446 + shift, Width = 200, Height = 18, Text = "Port B — Angle Control:" };

            // Port B live view
            liveBoxB = new RichTextBox
            {
                Left = 12,
                Top = 472 + shift,
                Width = 1080,
                Height = 160,
                ReadOnly = true,
                DetectUrls = false,
                WordWrap = false,
                Font = new Font("Consolas", 9)
            };

            Controls.AddRange(new Control[] {
        // Port A row
        portBox, baudBox, saveAsBtn, startBtn, stopBtn, clearABtn, statusLabel,
        // Port A send row + view
        sendBox, lineEndBox, sendBtn, echoCheck, aLabel, liveBoxA,
        // Charts
        speedChart, dirChart,
        // Port B row + view
        bHeader, port2Box, baud2Box, connect2Btn, disconnect2Btn, intBox,
        prefixLabel, prefixBox,
        sendModeBox, lineEnd2Box, sendIntBtn, echo2Check,
        bRecvLabel, clearBBtn, zeroBtn, liveBoxB
    });

            // Wire events — Port A
            saveAsBtn.Click += SaveAsBtn_Click;
            startBtn.Click += StartBtn_Click;
            stopBtn.Click += StopBtn_Click;
            clearABtn.Click += (s, e) => liveBoxA.Clear();

            sendBtn.Click += (s, e) => SendCurrentTextA();
            sendBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift && !e.Control && !e.Alt)
                {
                    e.Handled = true; e.SuppressKeyPress = true;
                    SendCurrentTextA();
                }
            };

            // Wire events — Port B
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
            clearBBtn.Click += (s, e) => liveBoxB.Clear();
            zeroBtn.Click += (s, e) => SendZeroOnB();
            sendModeBox.SelectedIndexChanged += SendModeBox_SelectedIndexChanged;

            SetSendUiEnabled(false); // Port A send disabled until connected
            //SetPortBUiEnabled(false); // Port B send/clear disabled until connected
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
            if (port2Box.Items.Count > 0) port2Box.SelectedIndex = Math.Min(1, port2Box.Items.Count - 1); // try a different one by default
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
            if (newFile) _writer.WriteLine("timestamp_iso,wind_speed,wind_direction,motor_angle_deg");
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

        private Chart CreateChart(string title, string xTitle, string yTitle)
        {
            var chart = new Chart();
            var area = new ChartArea();
            chart.ChartAreas.Add(area);

            // X axis as time
            area.AxisX.LabelStyle.Format = "HH:mm:ss";
            area.AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
            area.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

            area.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            area.AxisY.Title = yTitle;

            var series = new Series(title)
            {
                ChartType = SeriesChartType.FastLine,
                XValueType = ChartValueType.DateTime
            };
            chart.Series.Add(series);

            chart.Titles.Add(title);
            chart.Titles[0].Font = new Font("Segoe UI", 9, FontStyle.Bold);

            return chart;
        }

        // Replace your old AppendChartPoint(Series s, ...) with this version:
        private void AppendChartPoint(Chart chart, DateTime t, double value)
        {
            var s = chart.Series[0];
            s.Points.AddXY(t, value);
            if (s.Points.Count > MaxChartPoints)
                s.Points.RemoveAt(0);

            // Adjust X axis to keep latest window visible
            var area = chart.ChartAreas[0];
            if (s.Points.Count > 1)
            {
                double minX = s.Points[0].XValue;
                double maxX = s.Points[^1].XValue;
                area.AxisX.Minimum = minX;
                area.AxisX.Maximum = maxX;
            }
        }


        private void PortA_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string raw = _port.ReadLine();
                string line = raw.TrimEnd('\r', '\n', ' ');
                DateTime now = DateTime.Now;
                string ts = now.ToString("s");

                if (TryParseWind(line, out double ws, out double wd))
                {
                    // Build motor angle string once (blank if unknown)
                    double motor = Volatile.Read(ref _motorAngleDeg);
                    string motorAngleStr = double.IsNaN(motor)
                        ? ""
                        : motor.ToString(CultureInfo.InvariantCulture);

                    // CSV write
                    lock (this)
                    {
                        _writer.WriteLine($"{ts},{ws.ToString(CultureInfo.InvariantCulture)},{wd.ToString(CultureInfo.InvariantCulture)},{motorAngleStr}");
                        _writer.Flush();
                    }

                    // Live + charts
                    BeginInvoke(new Action(() =>
                    {
                        TrimBox(liveBoxA);
                        liveBoxA.AppendText($"{ts}  WS={ws}  WD={wd}\n");
                        liveBoxA.ScrollToCaret();

                        // Update charts
                        AppendChartPoint(speedChart, now, ws);
                        AppendChartPoint(dirChart, now, wd);
                        speedChart.Invalidate();
                        dirChart.Invalidate();
                    }));
                }
                else
                {
                    BeginInvoke(new Action(() =>
                    {
                        TrimBox(liveBoxA);
                        liveBoxA.AppendText($"{ts}  [A unparsed] {line}\n");
                        liveBoxA.ScrollToCaret();
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
            windSpeed = 0; windDirection = 0;
            if (string.IsNullOrWhiteSpace(line)) return false;

            // Expect: ",speed,direction"
            var parts = line.Split(',');
            if (parts.Length < 3) return false;

            var s = parts[1].Trim();
            var d = parts[2].Trim();

            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out windSpeed)
                && double.TryParse(d, NumberStyles.Float, CultureInfo.InvariantCulture, out windDirection);
        }


        // ----- Sending on Port A -----
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
                    liveBoxA.AppendText($"{ts}  >>A {text}{(suffix == "None" ? "" : $" ({suffix})")}\n");
                    liveBoxA.ScrollToCaret();
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

            // NEW: receive handler for Port B
            _port2.DataReceived += PortB_DataReceived;

            try
            {
                _port2.Open();
                statusLabel.Text = $"B: Connected on {_port2.PortName} @ {baud}";

                // Mark connected + set ASCII + refresh states
                SetPortBConnected(true);
                sendModeBox.SelectedIndex = 0;   // ensure ASCII for angles
                UpdatePortBControlsEnabled();

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
                    // NEW: prefix "A" or "R"
                    string prefix = (prefixBox.SelectedItem as string) ?? "A";

                    string suffix = lineEnd2Box.SelectedItem?.ToString() ?? "None";
                    string terminator = suffix switch
                    {
                        "\\n" => "\n",
                        "\\r" => "\r",
                        "\\r\\n" => "\r\n",
                        _ => string.Empty
                    };

                    string payload = prefix + value.ToString(CultureInfo.InvariantCulture) + terminator;
                    _port2.Write(payload);

                    if (echo2Check.Checked)
                    {
                        var ts = DateTime.Now.ToString("s");
                        string suffixNote = (suffix == "None") ? "" : $" ({suffix})";
                        liveBoxB.AppendText($"{ts}  >>B (ASCII) {prefix}{value}{suffixNote}\n");
                        liveBoxB.ScrollToCaret();
                    }
                }
                else // Single byte (0–255) — prefix not applicable
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
                        liveBoxB.AppendText($"{ts}  >>B (byte) 0x{b:X2} ({value})\n");
                        liveBoxB.ScrollToCaret();
                    }
                }
                // Keep focus in the angle box so Enter works repeatedly
                intBox.Focus();
                intBox.Select(0, intBox.Text.Length);   // highlight for quick overwrite
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

        // NEW: receive data from Port B and show in its own window
        private void PortB_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string chunk = _port2.ReadExisting();
                if (string.IsNullOrEmpty(chunk)) return;

                // Always mirror to the UI
                BeginInvoke(new Action(() =>
                {
                    TrimBox(liveBoxB);
                    liveBoxB.AppendText(chunk);
                    liveBoxB.ScrollToCaret();
                }));

                // Accumulate and search within a rolling window (no dependency on newlines)
                lock (_port2Buf)
                {
                    _port2Buf.Append(chunk);

                    // Normalize some common line endings to reduce weird splits (optional)
                    _port2Buf.Replace("\r\n", "\n");

                    // Scan the whole buffer for one or more "now at ..." occurrences
                    string buf = _port2Buf.ToString();
                    foreach (Match m in NowAtRegex.Matches(buf))
                    {
                        if (double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double ang))
                        {
                            Volatile.Write(ref _motorAngleDeg, ang);
                            BeginInvoke(new Action(() => statusLabel.Text = $"B: now at {ang:0.###}°"));
                        }
                    }

                    // Keep only the tail to avoid unbounded growth; 2 KB is plenty to cover split phrases
                    if (_port2Buf.Length > Port2BufMax)
                    {
                        _port2Buf.Remove(0, _port2Buf.Length - Port2BufMax);
                    }
                }
            }
            catch (TimeoutException) { }
            catch (IOException)
            {
                BeginInvoke(new Action(() => statusLabel.Text = "B: I/O error (device removed?)."));
            }
            catch (InvalidOperationException) { /* Port closed while reading */ }
        }



        // ----- Helpers -----

        private void SetPortBConnected(bool isConnected)
        {
            // Only these four are managed here:
            connect2Btn.Enabled = !isConnected;
            disconnect2Btn.Enabled = isConnected;
            port2Box.Enabled = !isConnected;
            baud2Box.Enabled = !isConnected;

            // Everything else (angle box, prefix, line ending, etc.)
            // is handled by UpdatePortBControlsEnabled():
            UpdatePortBControlsEnabled();
        }

        private void TrimBox(RichTextBox box)
        {
            const int maxChars = 500_000;
            if (box.TextLength > maxChars)
                box.Select(0, box.TextLength - maxChars);
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
                    _port2.DataReceived -= PortB_DataReceived;
                    if (_port2.IsOpen) _port2.Close();
                    _port2.Dispose();
                    SetPortBConnected(false);
                    statusLabel.Text = "B: Disconnected.";
                    _port2 = null;
                }
            }
            catch { }
            connect2Btn.Enabled = true;
            disconnect2Btn.Enabled = false;
            port2Box.Enabled = true;
            baud2Box.Enabled = true;
            //SetPortBUiEnabled(false);
        }

        private void SetSendUiEnabled(bool on)
        {
            sendBox.Enabled = on;
            sendBtn.Enabled = on;
            lineEndBox.Enabled = on;
            echoCheck.Enabled = on;
        }
        private void UpdatePortBControlsEnabled()
        {
            bool connected = _port2 != null && _port2.IsOpen;
            bool ascii = sendModeBox.SelectedIndex == 0;

            // Always enable the basics when connected
            intBox.Enabled = connected;
            sendIntBtn.Enabled = connected;
            sendModeBox.Enabled = connected;
            echo2Check.Enabled = connected;
            clearBBtn.Enabled = connected;
            zeroBtn.Enabled = connected;

            // Only in ASCII mode
            lineEnd2Box.Enabled = connected && ascii;
            prefixBox.Enabled = connected && ascii;
            //if (anticlockwiseCheck != null) anticlockwiseCheck.Enabled = connected && ascii;
        }

        // Call this on mode changes
        private void SendModeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePortBControlsEnabled();
        }


        private void SendZeroOnB()
        {
            try
            {
                if (_port2 == null || !_port2.IsOpen)
                {
                    System.Media.SystemSounds.Beep.Play();
                    statusLabel.Text = "B: Not connected — press Connect B first.";
                    return;
                }

                if (sendModeBox.SelectedIndex == 0) // ASCII mode
                {
                    string suffix = lineEnd2Box.SelectedItem?.ToString() ?? "None";
                    string terminator = suffix switch
                    {
                        "\\n" => "\n",
                        "\\r" => "\r",
                        "\\r\\n" => "\r\n",
                        _ => string.Empty
                    };

                    _port2.Write("Z" + terminator);

                    if (echo2Check.Checked)
                    {
                        var ts = DateTime.Now.ToString("s");
                        string suffixNote = (suffix == "None") ? "" : $" ({suffix})";
                        liveBoxB.AppendText($"{ts}  >>B Z{suffixNote}\n");
                        liveBoxB.ScrollToCaret();
                    }
                }
                else // Single byte mode
                {
                    byte z = 0x5A; // 'Z'
                    _port2.Write(new byte[] { z }, 0, 1);

                    if (echo2Check.Checked)
                    {
                        var ts = DateTime.Now.ToString("s");
                        liveBoxB.AppendText($"{ts}  >>B (byte) 0x5A ('Z')\n");
                        liveBoxB.ScrollToCaret();
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
    }
}
