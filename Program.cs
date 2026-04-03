using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

// Tray app MVP scaffold for:
// - enumerating monitors
// - enumerating USB devices
// - manual monitor input switching from tray/UI
// - optional auto-switch toggle
// - USB learning/setup flow (snapshot -> switch away -> compare -> switch back -> compare)
//
// Build target: Windows only, WinForms.
// Suggested csproj settings:
//   <TargetFramework>net8.0-windows</TargetFramework>
//   <UseWindowsForms>true</UseWindowsForms>
//
// This is an MVP scaffold with the main structure and working monitor enumeration / manual switching.
// USB learning is implemented as snapshot/diff logic using SetupAPI enumeration.
// Auto-switch toggle is wired in the UI but device-change event handling is intentionally left conservative.

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new TrayApplicationContext());
    }
}

internal sealed class TrayApplicationContext : ApplicationContext
{
    private readonly NotifyIcon _notifyIcon;
    private readonly MainForm _mainForm;

    public TrayApplicationContext()
    {
        _mainForm = new MainForm();
        _mainForm.FormClosing += MainForm_FormClosing;

        var menu = new ContextMenuStrip();
        menu.Items.Add("Open", null, (_, _) => ShowMainForm());
        menu.Items.Add("Refresh", null, (_, _) => _mainForm.RefreshAll());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) => ExitThread());

        _notifyIcon = new NotifyIcon
        {
            Text = "USB Monitor Switcher",
            Icon = SystemIcons.Application,
            Visible = true,
            ContextMenuStrip = menu
        };

        _notifyIcon.DoubleClick += (_, _) => ShowMainForm();
        ShowMainForm();
    }

    private void ShowMainForm()
    {
        if (_mainForm.Visible)
        {
            _mainForm.WindowState = FormWindowState.Normal;
            _mainForm.BringToFront();
            _mainForm.Activate();
            return;
        }

        _mainForm.Show();
        _mainForm.WindowState = FormWindowState.Normal;
        _mainForm.BringToFront();
        _mainForm.Activate();
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            _mainForm.Hide();
        }
    }

    protected override void ExitThreadCore()
    {
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _mainForm.Dispose();
        base.ExitThreadCore();
    }
}

internal sealed class MainForm : Form
{
    private readonly TabControl _tabs = new() { Dock = DockStyle.Fill };

    private readonly ListView _monitorList = CreateListView();
    private readonly ListView _usbList = CreateListView();
    private readonly TextBox _logBox = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
    private readonly CheckBox _autoSwitchCheck = new() { Text = "Enable auto-monitor-switching based on USB changes", Dock = DockStyle.Top, Height = 30 };

    private readonly Button _refreshMonitorsBtn = new() { Text = "Refresh monitors", AutoSize = true };
    private readonly Button _toPcBtn = new() { Text = "Switch selected monitor to PC", AutoSize = true };
    private readonly Button _toLaptopBtn = new() { Text = "Switch selected monitor to Laptop", AutoSize = true };

    private readonly Button _refreshUsbBtn = new() { Text = "Refresh USB devices", AutoSize = true };
    private readonly Button _snapshotBeforeBtn = new() { Text = "1) Capture current USB snapshot", AutoSize = true };
    private readonly Button _snapshotAwayBtn = new() { Text = "2) Switch USB away, then capture", AutoSize = true };
    private readonly Button _snapshotBackBtn = new() { Text = "3) Switch USB back, then capture", AutoSize = true };
    private readonly Button _proposeTrackedBtn = new() { Text = "Propose tracked USB devices", AutoSize = true };

    private readonly Label _wizardSummary = new() { Dock = DockStyle.Fill, AutoSize = false, TextAlign = ContentAlignment.TopLeft };

    private List<MonitorInfo> _monitors = new();
    private List<DeviceInfo> _usbDevices = new();

    private List<DeviceInfo> _snapshotBefore = new();
    private List<DeviceInfo> _snapshotAway = new();
    private List<DeviceInfo> _snapshotBack = new();

    public MainForm()
    {
        Text = "USB Monitor Switcher";
        Width = 1100;
        Height = 760;
        StartPosition = FormStartPosition.CenterScreen;

        BuildUi();
        WireEvents();
        RefreshAll();
    }

    public void RefreshAll()
    {
        RefreshMonitors();
        RefreshUsbDevices();
        Log("Refreshed monitors and USB devices.");
    }

    private void BuildUi()
    {
        Controls.Add(_tabs);

        // Status tab
        var statusTab = new TabPage("Status");
        var statusPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
        statusPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        statusPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        statusPanel.Controls.Add(_autoSwitchCheck, 0, 0);
        statusPanel.Controls.Add(_logBox, 0, 1);
        statusTab.Controls.Add(statusPanel);
        _tabs.TabPages.Add(statusTab);

        // Monitors tab
        var monitorsTab = new TabPage("Monitors");
        var monitorsPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
        monitorsPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        monitorsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var monitorButtons = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
        monitorButtons.Controls.AddRange(new Control[] { _refreshMonitorsBtn, _toPcBtn, _toLaptopBtn });
        monitorsPanel.Controls.Add(monitorButtons, 0, 0);
        monitorsPanel.Controls.Add(_monitorList, 0, 1);
        monitorsTab.Controls.Add(monitorsPanel);
        _tabs.TabPages.Add(monitorsTab);

        // USB devices tab
        var usbTab = new TabPage("USB Devices");
        var usbPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
        usbPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        usbPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var usbButtons = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
        usbButtons.Controls.Add(_refreshUsbBtn);
        usbPanel.Controls.Add(usbButtons, 0, 0);
        usbPanel.Controls.Add(_usbList, 0, 1);
        usbTab.Controls.Add(usbPanel);
        _tabs.TabPages.Add(usbTab);

        // Setup wizard tab
        var wizardTab = new TabPage("Setup Wizard");
        var wizardPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
        wizardPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        wizardPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var wizardButtons = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
        wizardButtons.Controls.AddRange(new Control[] { _snapshotBeforeBtn, _snapshotAwayBtn, _snapshotBackBtn, _proposeTrackedBtn });
        wizardPanel.Controls.Add(wizardButtons, 0, 0);
        wizardPanel.Controls.Add(_wizardSummary, 0, 1);
        wizardTab.Controls.Add(wizardPanel);
        _tabs.TabPages.Add(wizardTab);
    }

    private void WireEvents()
    {
        _refreshMonitorsBtn.Click += (_, _) => RefreshMonitors();
        _refreshUsbBtn.Click += (_, _) => RefreshUsbDevices();

        _toPcBtn.Click += (_, _) => ManualSwitchSelectedMonitor(toLaptop: false);
        _toLaptopBtn.Click += (_, _) => ManualSwitchSelectedMonitor(toLaptop: true);

        _snapshotBeforeBtn.Click += (_, _) =>
        {
            _snapshotBefore = DeviceEnumerator.GetPresentDevices().Where(IsLikelyUsbish).ToList();
            Log($"Captured current USB snapshot: {_snapshotBefore.Count} devices.");
            UpdateWizardSummary();
        };

        _snapshotAwayBtn.Click += (_, _) =>
        {
            _snapshotAway = DeviceEnumerator.GetPresentDevices().Where(IsLikelyUsbish).ToList();
            Log($"Captured 'away' USB snapshot: {_snapshotAway.Count} devices.");
            UpdateWizardSummary();
        };

        _snapshotBackBtn.Click += (_, _) =>
        {
            _snapshotBack = DeviceEnumerator.GetPresentDevices().Where(IsLikelyUsbish).ToList();
            Log($"Captured 'back' USB snapshot: {_snapshotBack.Count} devices.");
            UpdateWizardSummary();
        };

        _proposeTrackedBtn.Click += (_, _) =>
        {
            UpdateWizardSummary();
            _tabs.SelectedIndex = 3;
        };

        _autoSwitchCheck.CheckedChanged += (_, _) =>
        {
            Log(_autoSwitchCheck.Checked
                ? "Auto-switching enabled (event-driven detection not yet wired in this scaffold)."
                : "Auto-switching disabled.");
        };
    }

    private void RefreshMonitors()
    {
        _monitors = MonitorController.GetMonitors();
        _monitorList.Items.Clear();

        foreach (var m in _monitors)
        {
            var item = new ListViewItem(new[]
            {
                m.BestName,
                m.DeviceId,
                m.PhysicalDescription,
                m.IsAccessible ? "Yes" : "No"
            })
            {
                Tag = m
            };
            _monitorList.Items.Add(item);
        }

        EnsureMonitorColumns();
        Log($"Enumerated {_monitors.Count} monitor entries.");
    }

    private void RefreshUsbDevices()
    {
        _usbDevices = DeviceEnumerator.GetPresentDevices().Where(IsLikelyUsbish).ToList();
        _usbList.Items.Clear();

        foreach (var d in _usbDevices)
        {
            var item = new ListViewItem(new[]
            {
                d.FriendlyName,
                d.ClassName,
                d.InstanceId,
                d.LocationInfo,
                string.Join(" | ", d.LocationPaths.Take(2))
            })
            {
                Tag = d
            };
            _usbList.Items.Add(item);
        }

        EnsureUsbColumns();
        Log($"Enumerated {_usbDevices.Count} USB/HID-ish devices.");
    }

    private void ManualSwitchSelectedMonitor(bool toLaptop)
    {
        if (_monitorList.SelectedItems.Count == 0)
        {
            MessageBox.Show(this, "Select a monitor first.", "No monitor selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var selected = (MonitorInfo)_monitorList.SelectedItems[0].Tag!;
        var defaultValue = toLaptop ? 17u : 15u;
        var label = toLaptop ? "Laptop" : "PC";

        var input = Prompt.ShowDialog(this,
            $"Enter VCP 0x60 input value for '{selected.DisplayName}' -> {label}",
            "Manual monitor switch",
            defaultValue.ToString());

        if (string.IsNullOrWhiteSpace(input))
            return;

        if (!uint.TryParse(input, out var value))
        {
            MessageBox.Show(this, "Input value must be an unsigned integer.", "Invalid value", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var ok = MonitorController.TrySetInput(selected, value);
        Log(ok
            ? $"Switched '{selected.DisplayName}' to input value {value}."
            : $"Failed to switch '{selected.DisplayName}' to input value {value}.");
    }

    private void UpdateWizardSummary()
    {
        var beforeIds = new HashSet<string>(_snapshotBefore.Select(d => d.InstanceId), StringComparer.OrdinalIgnoreCase);
        var awayIds = new HashSet<string>(_snapshotAway.Select(d => d.InstanceId), StringComparer.OrdinalIgnoreCase);
        var backIds = new HashSet<string>(_snapshotBack.Select(d => d.InstanceId), StringComparer.OrdinalIgnoreCase);

        var disappearedWhenAway = _snapshotBefore
            .Where(d => !awayIds.Contains(d.InstanceId))
            .OrderBy(d => d.InstanceId)
            .ToList();

        var returnedWhenBack = _snapshotBack
            .Where(d => beforeIds.Contains(d.InstanceId) || disappearedWhenAway.Any(x => SameFamily(x, d)))
            .OrderBy(d => d.InstanceId)
            .ToList();

        var proposed = disappearedWhenAway
            .Where(d => d.InstanceId.Contains(@"USB\VID_", StringComparison.OrdinalIgnoreCase)
					|| d.InstanceId.Contains(@"HID\VID_", StringComparison.OrdinalIgnoreCase))
            .Select(d => FamilyKey(d.InstanceId))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x)
            .ToList();

        var sb = new StringBuilder();
        sb.AppendLine("Setup Wizard");
        sb.AppendLine();
        sb.AppendLine($"Snapshot before: {_snapshotBefore.Count}");
        sb.AppendLine($"Snapshot away  : {_snapshotAway.Count}");
        sb.AppendLine($"Snapshot back  : {_snapshotBack.Count}");
        sb.AppendLine();

        sb.AppendLine("Devices that disappeared when switched away from PC:");
        if (disappearedWhenAway.Count == 0)
            sb.AppendLine("  (none yet)");
        else
            foreach (var d in disappearedWhenAway.Take(40))
                sb.AppendLine($"  - {d.InstanceId}");

        sb.AppendLine();
        sb.AppendLine("Devices observed after switching back:");
        if (returnedWhenBack.Count == 0)
            sb.AppendLine("  (none yet)");
        else
            foreach (var d in returnedWhenBack.Take(40))
                sb.AppendLine($"  - {d.InstanceId}");

        sb.AppendLine();
        sb.AppendLine("Proposed tracked device families:");
        if (proposed.Count == 0)
            sb.AppendLine("  (none yet)");
        else
            foreach (var p in proposed)
                sb.AppendLine($"  - {p}");

        _wizardSummary.Text = sb.ToString();
    }

    private static bool SameFamily(DeviceInfo a, DeviceInfo b)
        => string.Equals(FamilyKey(a.InstanceId), FamilyKey(b.InstanceId), StringComparison.OrdinalIgnoreCase);

    private static string FamilyKey(string instanceId)
    {
        var upper = instanceId.ToUpperInvariant();
        var vid = ExtractToken(upper, "VID_");
        var pid = ExtractToken(upper, "PID_");

        if (!string.IsNullOrWhiteSpace(vid) && !string.IsNullOrWhiteSpace(pid))
            return $"VID_{vid}&PID_{pid}";

        return upper;
    }

    private static string ExtractToken(string input, string prefix)
    {
        var idx = input.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return string.Empty;
        idx += prefix.Length;
        if (idx + 4 > input.Length) return string.Empty;
        return input.Substring(idx, 4);
    }

    private static bool IsLikelyUsbish(DeviceInfo d)
		=> d.InstanceId.Contains(@"USB\", StringComparison.OrdinalIgnoreCase)
		|| d.InstanceId.Contains(@"HID\", StringComparison.OrdinalIgnoreCase)
        || d.EnumeratorName.Contains("USB", StringComparison.OrdinalIgnoreCase)
        || d.ClassName.Contains("HID", StringComparison.OrdinalIgnoreCase)
        || d.ClassName.Contains("Keyboard", StringComparison.OrdinalIgnoreCase)
        || d.ClassName.Contains("Mouse", StringComparison.OrdinalIgnoreCase);

    private void Log(string message)
    {
        _logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
    }

    private static ListView CreateListView()
    {
        return new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false
        };
    }

    private void EnsureMonitorColumns()
    {
        if (_monitorList.Columns.Count > 0) return;
        _monitorList.Columns.Add("Monitor name", 220);
        _monitorList.Columns.Add("Device ID", 260);
        _monitorList.Columns.Add("Physical description", 320);
        _monitorList.Columns.Add("DDC accessible", 100);
    }

    private void EnsureUsbColumns()
    {
        if (_usbList.Columns.Count > 0) return;
        _usbList.Columns.Add("Friendly name", 220);
        _usbList.Columns.Add("Class", 120);
        _usbList.Columns.Add("Instance ID", 340);
        _usbList.Columns.Add("Location info", 180);
        _usbList.Columns.Add("Location paths", 320);
    }
}

internal static class Prompt
{
    public static string ShowDialog(IWin32Window owner, string text, string caption, string defaultValue)
    {
        using var form = new Form { Width = 560, Height = 160, Text = caption, StartPosition = FormStartPosition.CenterParent };
        var label = new Label { Left = 12, Top = 12, Width = 520, Text = text };
        var textBox = new TextBox { Left = 12, Top = 40, Width = 520, Text = defaultValue };
        var ok = new Button { Text = "OK", Left = 376, Width = 75, Top = 76, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Cancel", Left = 457, Width = 75, Top = 76, DialogResult = DialogResult.Cancel };
        form.Controls.AddRange(new Control[] { label, textBox, ok, cancel });
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        return form.ShowDialog(owner) == DialogResult.OK ? textBox.Text : string.Empty;
    }
}

internal sealed class MonitorInfo
{
    public string DisplayName { get; init; } = "";
    public string MonitorDeviceName { get; init; } = "";
    public string PhysicalDescription { get; init; } = "";
    public string DeviceId { get; init; } = "";
    public string DeviceString { get; init; } = "";
    public string EdidName { get; init; } = "";
    public IntPtr HMonitor { get; init; }
    public IntPtr PhysicalHandle { get; init; }
    public bool IsAccessible { get; init; }

    public string BestName
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(EdidName))
                return EdidName;

            if (!string.IsNullOrWhiteSpace(DeviceString) &&
                !string.Equals(DeviceString, "Generic PnP Monitor", StringComparison.OrdinalIgnoreCase))
                return DeviceString;

            if (!string.IsNullOrWhiteSpace(PhysicalDescription) &&
                !string.Equals(PhysicalDescription, "Generic PnP Monitor", StringComparison.OrdinalIgnoreCase))
                return PhysicalDescription;

            var parsed = MonitorNameResolver.TryGetFallbackNameFromDeviceId(DeviceId);
            if (!string.IsNullOrWhiteSpace(parsed))
                return parsed;

            if (!string.IsNullOrWhiteSpace(MonitorDeviceName))
                return MonitorDeviceName;

            return DisplayName;
        }
    }
}

internal static class MonitorNameResolver
{
    public static (string deviceId, string deviceString) GetMonitorDeviceForDisplay(string displayName)
    {
        uint devNum = 0;
        while (true)
        {
            var dd = new NativeMethods.DISPLAY_DEVICE();
            dd.cb = Marshal.SizeOf<NativeMethods.DISPLAY_DEVICE>();

            if (!NativeMethods.EnumDisplayDevices(displayName, devNum, ref dd, 0))
                break;

            if (!string.IsNullOrWhiteSpace(dd.DeviceID))
                return (dd.DeviceID, dd.DeviceString ?? string.Empty);

            devNum++;
        }

        return (string.Empty, string.Empty);
    }

    public static string TryGetFallbackNameFromDeviceId(string deviceId)
    {
        if (string.IsNullOrWhiteSpace(deviceId))
            return string.Empty;

        // Expected format often looks like:
        // MONITOR\ACI27A1\{...}
        var parts = deviceId.Split('\\');
        if (parts.Length < 2)
            return string.Empty;

        return parts[1];
    }
}

internal static class EdidParser
{
    public static string? TryGetMonitorName(byte[] edid)
    {
        if (edid == null || edid.Length < 128)
            return null;

        for (int offset = 54; offset <= 108; offset += 18)
        {
            if (edid[offset] == 0x00 &&
                edid[offset + 1] == 0x00 &&
                edid[offset + 2] == 0x00 &&
                edid[offset + 3] == 0xFC &&
                edid[offset + 4] == 0x00)
            {
                var nameBytes = new byte[13];
                Array.Copy(edid, offset + 5, nameBytes, 0, 13);

                var name = Encoding.ASCII.GetString(nameBytes)
                    .TrimEnd('\0', ' ', '\r', '\n');

                return string.IsNullOrWhiteSpace(name) ? null : name;
            }
        }

        return null;
    }
}

internal static class MonitorRegistryReader
{
    public static string? TryGetEdidName(string deviceId)
    {
        if (string.IsNullOrWhiteSpace(deviceId))
            return null;

        var classGuid = NativeMethods.GUID_DEVCLASS_MONITOR;

        IntPtr infoSet = NativeMethods.SetupDiGetClassDevs(
            ref classGuid,
            null,
            IntPtr.Zero,
            NativeMethods.DIGCF_PRESENT);

        if (infoSet == NativeMethods.INVALID_HANDLE_VALUE)
            return null;

        try
        {
            uint index = 0;
            while (true)
            {
                var devInfo = new NativeMethods.SP_DEVINFO_DATA();
                devInfo.cbSize = (uint)Marshal.SizeOf<NativeMethods.SP_DEVINFO_DATA>();

                if (!NativeMethods.SetupDiEnumDeviceInfo(infoSet, index, ref devInfo))
                {
                    int err = Marshal.GetLastWin32Error();
                    if (err == NativeMethods.ERROR_NO_MORE_ITEMS)
                        break;

                    return null;
                }

                index++;

                var instanceId = GetDeviceInstanceId(infoSet, ref devInfo);
                if (string.IsNullOrWhiteSpace(instanceId))
                    continue;

                if (!IsSameMonitorHardware(instanceId, deviceId))
					continue;

                IntPtr hKey = NativeMethods.SetupDiOpenDevRegKey(
                    infoSet,
                    ref devInfo,
                    NativeMethods.DICS_FLAG_GLOBAL,
                    0,
                    NativeMethods.DIREG_DEV,
                    NativeMethods.KEY_READ);

                if (hKey == NativeMethods.INVALID_HANDLE_VALUE || hKey == IntPtr.Zero)
                    return null;

                try
                {
                    var edid = ReadRegistryBinaryValue(hKey, "EDID");
                    if (edid == null || edid.Length == 0)
                        return null;

                    return EdidParser.TryGetMonitorName(edid);
                }
                finally
                {
                    NativeMethods.RegCloseKey(hKey);
                }
            }
        }
        finally
        {
            NativeMethods.SetupDiDestroyDeviceInfoList(infoSet);
        }

        return null;
    }

    private static string? GetDeviceInstanceId(IntPtr infoSet, ref NativeMethods.SP_DEVINFO_DATA devInfo)
    {
        var sb = new StringBuilder(512);
        return NativeMethods.SetupDiGetDeviceInstanceId(infoSet, ref devInfo, sb, sb.Capacity, out _)
            ? sb.ToString()
            : null;
    }

    private static byte[]? ReadRegistryBinaryValue(IntPtr hKey, string valueName)
    {
        uint type;
        uint size = 0;

        int rc = NativeMethods.RegQueryValueEx(
            hKey,
            valueName,
            IntPtr.Zero,
            out type,
            null,
            ref size);

        if (rc != 0 || size == 0 || type != NativeMethods.REG_BINARY)
            return null;

        byte[] data = new byte[size];
        rc = NativeMethods.RegQueryValueEx(
            hKey,
            valueName,
            IntPtr.Zero,
            out type,
            data,
            ref size);

        if (rc != 0 || type != NativeMethods.REG_BINARY)
            return null;

        return data;
    }
	
	private static bool IsSameMonitorHardware(string instanceId, string deviceId)
{
    var a = GetMonitorHardwareKey(instanceId);
    var b = GetMonitorHardwareKey(deviceId);

    return !string.IsNullOrWhiteSpace(a) &&
           !string.IsNullOrWhiteSpace(b) &&
           string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
}

private static string GetMonitorHardwareKey(string value)
{
    if (string.IsNullOrWhiteSpace(value))
        return string.Empty;

    var parts = value.Split('\\');
    if (parts.Length < 2)
        return string.Empty;

    return $@"{parts[0]}\{parts[1]}";
}
}

internal static class MonitorController
{
    public static List<MonitorInfo> GetMonitors()
    {
        var result = new List<MonitorInfo>();

        NativeMethods.EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, (hMonitor, _, _, _) =>
        {
            var mi = new NativeMethods.MONITORINFOEX();
            mi.cbSize = Marshal.SizeOf<NativeMethods.MONITORINFOEX>();
            NativeMethods.GetMonitorInfo(hMonitor, ref mi);

            var (deviceId, deviceString) = MonitorNameResolver.GetMonitorDeviceForDisplay(mi.szDevice);
            var edidName = MonitorRegistryReader.TryGetEdidName(deviceId) ?? string.Empty;

            if (NativeMethods.GetNumberOfPhysicalMonitorsFromHMONITOR(hMonitor, out uint count) && count > 0)
            {
                var arr = new NativeMethods.PHYSICAL_MONITOR[count];
                if (NativeMethods.GetPhysicalMonitorsFromHMONITOR(hMonitor, count, arr))
                {
                    foreach (var pm in arr)
                    {
                        result.Add(new MonitorInfo
                        {
                            DisplayName = mi.szDevice,
                            MonitorDeviceName = mi.szDevice,
                            PhysicalDescription = pm.szPhysicalMonitorDescription,
                            DeviceId = deviceId,
                            DeviceString = deviceString,
                            EdidName = edidName,
                            HMonitor = hMonitor,
                            PhysicalHandle = pm.hPhysicalMonitor,
                            IsAccessible = pm.hPhysicalMonitor != IntPtr.Zero,
                        });
                    }
					System.Diagnostics.Debug.WriteLine(
					$"Display={mi.szDevice}, DeviceId={deviceId}, EdidName={edidName}");
                }
            }

            return true;
        }, IntPtr.Zero);

        return result;
    }

    public static bool TrySetInput(MonitorInfo monitor, uint inputValue)
    {
        if (monitor.PhysicalHandle == IntPtr.Zero) return false;
        return NativeMethods.SetVCPFeature(monitor.PhysicalHandle, 0x60, inputValue);
    }
}

internal sealed class DeviceInfo
{
    public string FriendlyName { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string EnumeratorName { get; init; } = "";
    public string InstanceId { get; init; } = "";
    public string LocationInfo { get; init; } = "";
    public List<string> LocationPaths { get; init; } = new();
}

internal static class DeviceEnumerator
{
    public static List<DeviceInfo> GetPresentDevices()
    {
        var result = new List<DeviceInfo>();
        var classGuid = Guid.Empty;

        IntPtr infoSet = NativeMethods.SetupDiGetClassDevs(ref classGuid, null, IntPtr.Zero,
            NativeMethods.DIGCF_ALLCLASSES | NativeMethods.DIGCF_PRESENT);

        if (infoSet == NativeMethods.INVALID_HANDLE_VALUE)
            throw new InvalidOperationException("SetupDiGetClassDevs failed.");

        try
        {
            uint index = 0;
            while (true)
            {
                var devInfo = new NativeMethods.SP_DEVINFO_DATA();
                devInfo.cbSize = (uint)Marshal.SizeOf<NativeMethods.SP_DEVINFO_DATA>();

                if (!NativeMethods.SetupDiEnumDeviceInfo(infoSet, index, ref devInfo))
                {
                    var err = Marshal.GetLastWin32Error();
                    if (err == NativeMethods.ERROR_NO_MORE_ITEMS) break;
                    throw new InvalidOperationException($"SetupDiEnumDeviceInfo failed: {err}");
                }
                index++;

                result.Add(new DeviceInfo
                {
                    FriendlyName = GetDeviceDisplayName(infoSet, ref devInfo)
                        ?? GetStringProp(infoSet, ref devInfo, NativeMethods.SPDRP_FRIENDLYNAME)
                        ?? GetStringProp(infoSet, ref devInfo, NativeMethods.SPDRP_DEVICEDESC)
                        ?? "<unknown>",
                    ClassName = GetStringProp(infoSet, ref devInfo, NativeMethods.SPDRP_CLASS) ?? "",
                    EnumeratorName = GetStringProp(infoSet, ref devInfo, NativeMethods.SPDRP_ENUMERATOR_NAME) ?? "",
                    InstanceId = GetDeviceInstanceId(infoSet, ref devInfo) ?? "",
                    LocationInfo = GetStringProp(infoSet, ref devInfo, NativeMethods.SPDRP_LOCATION_INFORMATION) ?? "",
                    LocationPaths = GetMultiSzProp(infoSet, ref devInfo, NativeMethods.SPDRP_LOCATION_PATHS),
                });
            }
        }
        finally
        {
            NativeMethods.SetupDiDestroyDeviceInfoList(infoSet);
        }

        return result;
    }

    private static string? GetDeviceInstanceId(IntPtr infoSet, ref NativeMethods.SP_DEVINFO_DATA devInfo)
    {
        var sb = new StringBuilder(1024);
        return NativeMethods.SetupDiGetDeviceInstanceId(infoSet, ref devInfo, sb, sb.Capacity, out _)
            ? sb.ToString()
            : null;
    }

    private static string? GetDeviceDisplayName(IntPtr infoSet, ref NativeMethods.SP_DEVINFO_DATA devInfo)
    {
        if (!TryGetDevicePropertyString(infoSet, ref devInfo, NativeMethods.DEVPKEY_NAME, out var value))
            return null;

        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static string? GetStringProp(IntPtr infoSet, ref NativeMethods.SP_DEVINFO_DATA devInfo, uint prop)
    {
        if (!TryGetRegistryProperty(infoSet, ref devInfo, prop, out var bytes))
            return null;

        var s = Encoding.Unicode.GetString(bytes).TrimEnd('\0');
        return string.IsNullOrWhiteSpace(s) ? null : s;
    }

    private static List<string> GetMultiSzProp(IntPtr infoSet, ref NativeMethods.SP_DEVINFO_DATA devInfo, uint prop)
    {
        if (!TryGetRegistryProperty(infoSet, ref devInfo, prop, out var bytes))
            return new List<string>();

        var s = Encoding.Unicode.GetString(bytes).TrimEnd('\0');
			return s.Split('\0', StringSplitOptions.RemoveEmptyEntries).ToList();
    }

    private static bool TryGetDevicePropertyString(IntPtr infoSet, ref NativeMethods.SP_DEVINFO_DATA devInfo, NativeMethods.DEVPROPKEY propertyKey, out string value)
    {
        value = string.Empty;

        NativeMethods.SetupDiGetDevicePropertyW(infoSet, ref devInfo, ref propertyKey, out _, IntPtr.Zero, 0, out var requiredSize, 0);
        if (requiredSize == 0)
            return false;

        var buffer = Marshal.AllocHGlobal((int)requiredSize);
        try
        {
            if (!NativeMethods.SetupDiGetDevicePropertyW(infoSet, ref devInfo, ref propertyKey, out _, buffer, requiredSize, out _, 0))
                return false;

            value = Marshal.PtrToStringUni(buffer) ?? string.Empty;
            return !string.IsNullOrWhiteSpace(value);
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static bool TryGetRegistryProperty(IntPtr infoSet, ref NativeMethods.SP_DEVINFO_DATA devInfo, uint property, out byte[] data)
    {
        data = Array.Empty<byte>();
        NativeMethods.SetupDiGetDeviceRegistryProperty(infoSet, ref devInfo, property, out _, IntPtr.Zero, 0, out var size);
        if (size == 0) return false;

        var buffer = new byte[size];
        if (!NativeMethods.SetupDiGetDeviceRegistryProperty(infoSet, ref devInfo, property, out _, buffer, (uint)buffer.Length, out _))
            return false;

        data = buffer;
        return true;
    }
}

internal static class NativeMethods
{
    public static readonly IntPtr INVALID_HANDLE_VALUE = new(-1);
    public const int ERROR_NO_MORE_ITEMS = 259;

    public const uint DIGCF_PRESENT = 0x00000002;
    public const uint DIGCF_ALLCLASSES = 0x00000004;

    public const uint SPDRP_DEVICEDESC = 0x00000000;
    public const uint SPDRP_CLASS = 0x00000007;
    public const uint SPDRP_FRIENDLYNAME = 0x0000000C;
    public const uint SPDRP_LOCATION_INFORMATION = 0x0000000D;
    public const uint SPDRP_ENUMERATOR_NAME = 0x00000016;
    public const uint SPDRP_LOCATION_PATHS = 0x00000023;

    public const uint DICS_FLAG_GLOBAL = 0x00000001;
    public const uint DIREG_DEV = 0x00000001;
    public const uint KEY_READ = 0x20019;
    public const uint REG_BINARY = 3;

    public static readonly DEVPROPKEY DEVPKEY_NAME =
        new(new Guid("b725f130-47ef-101a-a5f1-02608c9eebac"), 10);

    public static readonly Guid GUID_DEVCLASS_MONITOR =
        new("4d36e96e-e325-11ce-bfc1-08002be10318");

    [StructLayout(LayoutKind.Sequential)]
    public struct DEVPROPKEY
    {
        public Guid fmtid;
        public uint pid;

        public DEVPROPKEY(Guid fmtid, uint pid)
        {
            this.fmtid = fmtid;
            this.pid = pid;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SP_DEVINFO_DATA
    {
        public uint cbSize;
        public Guid ClassGuid;
        public uint DevInst;
        public IntPtr Reserved;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct PHYSICAL_MONITOR
    {
        public IntPtr hPhysicalMonitor;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string szPhysicalMonitorDescription;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DISPLAY_DEVICE
    {
        public int cb;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string DeviceName;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceString;

        public int StateFlags;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceID;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceKey;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct MONITORINFOEX
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string szDevice;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    public delegate bool MonitorEnumProc(
        IntPtr hMonitor,
        IntPtr hdcMonitor,
        IntPtr lprcMonitor,
        IntPtr dwData);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr SetupDiGetClassDevs(
        ref Guid ClassGuid,
        string? Enumerator,
        IntPtr hwndParent,
        uint Flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    public static extern bool SetupDiEnumDeviceInfo(
        IntPtr DeviceInfoSet,
        uint MemberIndex,
        ref SP_DEVINFO_DATA DeviceInfoData);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool SetupDiGetDeviceInstanceId(
        IntPtr DeviceInfoSet,
        ref SP_DEVINFO_DATA DeviceInfoData,
        StringBuilder DeviceInstanceId,
        int DeviceInstanceIdSize,
        out int RequiredSize);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SetupDiGetDeviceRegistryPropertyW")]
    public static extern bool SetupDiGetDeviceRegistryProperty(
        IntPtr DeviceInfoSet,
        ref SP_DEVINFO_DATA DeviceInfoData,
        uint Property,
        out uint PropertyRegDataType,
        byte[] PropertyBuffer,
        uint PropertyBufferSize,
        out uint RequiredSize);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SetupDiGetDeviceRegistryPropertyW")]
    public static extern bool SetupDiGetDeviceRegistryProperty(
        IntPtr DeviceInfoSet,
        ref SP_DEVINFO_DATA DeviceInfoData,
        uint Property,
        out uint PropertyRegDataType,
        IntPtr PropertyBuffer,
        uint PropertyBufferSize,
        out uint RequiredSize);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SetupDiGetDevicePropertyW")]
    public static extern bool SetupDiGetDevicePropertyW(
        IntPtr DeviceInfoSet,
        ref SP_DEVINFO_DATA DeviceInfoData,
        ref DEVPROPKEY PropertyKey,
        out uint PropertyType,
        IntPtr PropertyBuffer,
        uint PropertyBufferSize,
        out uint RequiredSize,
        uint Flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    public static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

    [DllImport("setupapi.dll", SetLastError = true)]
    public static extern IntPtr SetupDiOpenDevRegKey(
        IntPtr DeviceInfoSet,
        ref SP_DEVINFO_DATA DeviceInfoData,
        uint Scope,
        uint HwProfile,
        uint KeyType,
        uint samDesired);

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern int RegQueryValueEx(
        IntPtr hKey,
        string lpValueName,
        IntPtr lpReserved,
        out uint lpType,
        byte[]? lpData,
        ref uint lpcbData);

    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern int RegCloseKey(IntPtr hKey);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumDisplayMonitors(
        IntPtr hdc,
        IntPtr lprcClip,
        MonitorEnumProc lpfnEnum,
        IntPtr dwData);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool GetMonitorInfo(
        IntPtr hMonitor,
        ref MONITORINFOEX lpmi);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool EnumDisplayDevices(
        string? lpDevice,
        uint iDevNum,
        ref DISPLAY_DEVICE lpDisplayDevice,
        uint dwFlags);

    [DllImport("dxva2.dll", SetLastError = true)]
    public static extern bool GetNumberOfPhysicalMonitorsFromHMONITOR(
        IntPtr hMonitor,
        out uint pdwNumberOfPhysicalMonitors);

    [DllImport("dxva2.dll", SetLastError = true)]
    public static extern bool GetPhysicalMonitorsFromHMONITOR(
        IntPtr hMonitor,
        uint dwPhysicalMonitorArraySize,
        [Out] PHYSICAL_MONITOR[] pPhysicalMonitorArray);

    [DllImport("dxva2.dll", SetLastError = true)]
    public static extern bool SetVCPFeature(
        IntPtr hMonitor,
        byte bVCPCode,
        uint dwNewValue);
}
