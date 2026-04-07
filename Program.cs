using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.Json;

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

internal sealed class AppConfig
{
    public List<string> TrackedFamilies { get; set; } = new();
    public uint PcInputValue { get; set; } = 15;
    public uint LaptopInputValue { get; set; } = 17;
    public bool AutoSwitchEnabled { get; set; } = false;
    public List<string> SelectedMonitorIds { get; set; } = new();
}

internal sealed class MainForm : Form
{
    private enum SwitchState
    {
        Unknown,
        OnPc,
        Away
    }

    private readonly TabControl _tabs = new() { Dock = DockStyle.Fill };
    private readonly string _configDirectory = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
    "USBSwitch");

	private readonly string _configPath;

    private AppConfig _config = new();

    private readonly ListView _monitorList = new()
    {
        Dock = DockStyle.Fill,
        View = View.Details,
        FullRowSelect = true,
        GridLines = true,
        MultiSelect = false,
        CheckBoxes = true
    };
    private readonly ListView _usbList = CreateListView();
    private readonly TextBox _logBox = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
    private readonly CheckBox _autoSwitchCheck = new() { Text = "Enable auto-monitor-switching based on USB changes", Dock = DockStyle.Top, Height = 30 };

    private readonly Button _refreshMonitorsBtn = new() { Text = "Refresh monitors", AutoSize = true };
    private readonly Button _toPcBtn = new() { Text = "Switch selected monitor to PC", AutoSize = true };
    private readonly Button _toLaptopBtn = new() { Text = "Switch selected monitor to Laptop", AutoSize = true };

    private readonly Button _refreshUsbBtn = new() { Text = "Refresh USB devices", AutoSize = true };
    private readonly Button _guidedWizardBtn = new() { Text = "Run guided USB wizard (10s)", AutoSize = true };
    private readonly Button _snapshotBeforeBtn = new() { Text = "1) Capture current USB snapshot", AutoSize = true };
    private readonly Button _snapshotAwayBtn = new() { Text = "2) Switch USB away, then capture", AutoSize = true };
    private readonly Button _snapshotBackBtn = new() { Text = "3) Switch USB back, then capture", AutoSize = true };
    private readonly Button _proposeTrackedBtn = new() { Text = "Propose tracked USB devices", AutoSize = true };

    private readonly Label _wizardSummary = new() { Dock = DockStyle.Fill, AutoSize = false, TextAlign = ContentAlignment.TopLeft };
    private readonly System.Windows.Forms.Timer _wizardTimer = new() { Interval = 10_000 };
    private readonly System.Windows.Forms.Timer _usbDebounceTimer = new() { Interval = 1500 };
    private readonly System.Windows.Forms.Timer _cooldownTimer = new() { Interval = 8000 };
	private readonly System.Windows.Forms.Timer _startupRefreshTimer = new() { Interval = 1500 };

    private List<MonitorInfo> _monitors = new();
    private List<DeviceInfo> _usbDevices = new();

    private List<DeviceInfo> _snapshotBefore = new();
    private List<DeviceInfo> _snapshotAway = new();
    private List<DeviceInfo> _snapshotBack = new();

    private int _guidedWizardPhase = 0;

    private HashSet<string> _trackedFamilies = new(StringComparer.OrdinalIgnoreCase);
    private SwitchState _switchState = SwitchState.Unknown;
    private bool _isInCooldown = false;
    private bool _updatingMonitorChecks;

public MainForm()
{
    Text = "USB Monitor Switcher";
    Width = 1100;
    Height = 760;
    StartPosition = FormStartPosition.CenterScreen;
	_configPath = Path.Combine(_configDirectory, "usb-monitor-switcher.json");
	Directory.CreateDirectory(_configDirectory);
	
    LoadConfig();
	_config.SelectedMonitorIds ??= new List<string>();
    _trackedFamilies = (_config.TrackedFamilies ?? new List<string>())
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    BuildUi();
    WireEvents();

    _autoSwitchCheck.Checked = _config.AutoSwitchEnabled;

    RefreshAll();

    Log(_trackedFamilies.Count > 0
        ? $"Loaded {_trackedFamilies.Count} tracked families from config."
        : "No tracked families loaded from config.");

    Log(_autoSwitchCheck.Checked
        ? "Auto-switching restored as enabled from config."
        : "Auto-switching restored as disabled from config.");


Log(_autoSwitchCheck.Checked
    ? "Auto-switching restored as enabled from config."
    : "Auto-switching restored as disabled from config.");

        if (_trackedFamilies.Count > 0)
            Log($"Loaded {_trackedFamilies.Count} tracked families from config.");
    }

    protected override void OnShown(EventArgs e)
{
    base.OnShown(e);
    _startupRefreshTimer.Start();
}
	
	private void StartupRefreshTimer_Tick(object? sender, EventArgs e)
{
    _startupRefreshTimer.Stop();
    RefreshMonitors();
    Log("Performed delayed startup monitor refresh.");
}
	
	protected override void WndProc(ref Message m)
    {
        const int WM_DEVICECHANGE = 0x0219;
        const int DBT_DEVICEARRIVAL = 0x8000;
        const int DBT_DEVICEREMOVECOMPLETE = 0x8004;
        const int DBT_DEVNODES_CHANGED = 0x0007;

        if (m.Msg == WM_DEVICECHANGE)
        {
            int eventType = m.WParam.ToInt32();
            if (eventType == DBT_DEVICEARRIVAL ||
                eventType == DBT_DEVICEREMOVECOMPLETE ||
                eventType == DBT_DEVNODES_CHANGED)
            {
                OnUsbTopologyChanged();
            }
        }

        base.WndProc(ref m);
    }

    protected override void Dispose(bool disposing)
{
    if (disposing)
    {
        ReleaseMonitorHandles();
        _wizardTimer.Dispose();
        _usbDebounceTimer.Dispose();
        _cooldownTimer.Dispose();
        _startupRefreshTimer.Dispose();
    }

    base.Dispose(disposing);
}

    public void RefreshAll()
    {
        RefreshMonitors();
        RefreshUsbDevices();
        Log("Refreshed monitors and USB devices.");
    }

	private void LoadConfig()
{
    try
    {
        if (!File.Exists(_configPath))
        {
            _config = new AppConfig();
            return;
        }

        var json = File.ReadAllText(_configPath);
		_config = JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
		_config.TrackedFamilies ??= new List<string>();
		_config.SelectedMonitorIds ??= new List<string>();
        _config.SelectedMonitorIds ??= new List<string>();

        Log($"Config loaded: {_configPath}");
    }
    catch (Exception ex)
    {
        _config = new AppConfig();
        Log($"Failed to load config: {ex}");
    }
}

    private void SaveConfig()
{
    try
    {
        Directory.CreateDirectory(_configDirectory);

        var json = JsonSerializer.Serialize(_config, new JsonSerializerOptions
        {
            WriteIndented = true
        });

        File.WriteAllText(_configPath, json);
        Log($"Config saved: {_configPath}");
    }
    catch (Exception ex)
    {
        Log($"Failed to save config: {ex}");
    }
}


private string GetMonitorSelectionKey(MonitorInfo monitor)
{
    if (monitor == null)
        return string.Empty;

    if (!string.IsNullOrWhiteSpace(monitor.DeviceId))
        return monitor.DeviceId;

    return monitor.DisplayName ?? string.Empty;
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
        wizardButtons.Controls.AddRange(new Control[]
        {
            _guidedWizardBtn,
            _snapshotBeforeBtn,
            _snapshotAwayBtn,
            _snapshotBackBtn,
            _proposeTrackedBtn
        });
        wizardPanel.Controls.Add(wizardButtons, 0, 0);
        wizardPanel.Controls.Add(_wizardSummary, 0, 1);
        wizardTab.Controls.Add(wizardPanel);
        _tabs.TabPages.Add(wizardTab);
    }

    private void WireEvents()
    {
        _refreshMonitorsBtn.Click += (_, _) => RefreshMonitors();
        _monitorList.ItemChecked += MonitorList_ItemChecked;
        _refreshUsbBtn.Click += (_, _) => RefreshUsbDevices();
		
		_startupRefreshTimer.Tick += StartupRefreshTimer_Tick;

        _toPcBtn.Click += (_, _) => ManualSwitchSelectedMonitor(toLaptop: false);
        _toLaptopBtn.Click += (_, _) => ManualSwitchSelectedMonitor(toLaptop: true);

        _guidedWizardBtn.Click += (_, _) => StartGuidedWizard();
        _wizardTimer.Tick += WizardTimer_Tick;

        _usbDebounceTimer.Tick += UsbDebounceTimer_Tick;
        _cooldownTimer.Tick += CooldownTimer_Tick;

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
    _config.AutoSwitchEnabled = _autoSwitchCheck.Checked;
    SaveConfig();

    if (_autoSwitchCheck.Checked)
    {
        Log(_trackedFamilies.Count > 0
            ? $"Auto-switching enabled with {_trackedFamilies.Count} tracked families."
            : "Auto-switching enabled, but no tracked families have been learned yet.");
    }
    else
    {
        Log("Auto-switching disabled.");
    }
};
    }


private void MonitorList_ItemChecked(object? sender, ItemCheckedEventArgs e)
{
    if (_updatingMonitorChecks)
        return;

    if (_config == null)
        return;

    var selected = new List<string>();

    foreach (ListViewItem item in _monitorList.Items)
    {
        if (item == null || !item.Checked)
            continue;

        if (item.Tag is not MonitorInfo monitor)
            continue;

        var key = GetMonitorSelectionKey(monitor);
        if (string.IsNullOrWhiteSpace(key))
            continue;

        selected.Add(key);
    }

    _config.SelectedMonitorIds = selected;
    SaveConfig();

    Log($"Saved {selected.Count} selected monitor(s) for auto-switch.");
}

    private void StartGuidedWizard()
    {
        if (_guidedWizardPhase != 0)
        {
            Log("Guided wizard is already running.");
            return;
        }

        _snapshotBefore = DeviceEnumerator.GetPresentDevices().Where(IsLikelyUsbish).ToList();
        _snapshotAway = new List<DeviceInfo>();
        _snapshotBack = new List<DeviceInfo>();

        _guidedWizardPhase = 1;
        _tabs.SelectedIndex = 3;

        _wizardSummary.Text =
            "Guided USB Wizard\r\n\r\n" +
            "Step 1 complete: captured current USB snapshot.\r\n\r\n" +
            "Now switch the USB switch to the laptop.\r\n" +
            "The app will capture the 'away' snapshot automatically in 10 seconds.";

        Log($"Guided wizard: captured current USB snapshot: {_snapshotBefore.Count} devices.");
        Log("Guided wizard: switch the USB switch to the laptop now. Capturing 'away' snapshot in 10 seconds.");

        _wizardTimer.Start();
    }

    private void WizardTimer_Tick(object? sender, EventArgs e)
    {
        _wizardTimer.Stop();

        if (_guidedWizardPhase == 1)
        {
            _snapshotAway = DeviceEnumerator.GetPresentDevices().Where(IsLikelyUsbish).ToList();
            _guidedWizardPhase = 2;

            Log($"Guided wizard: captured 'away' USB snapshot: {_snapshotAway.Count} devices.");
            _wizardSummary.Text =
                "Guided USB Wizard\r\n\r\n" +
                "Step 2 complete: captured snapshot while switched away.\r\n\r\n" +
                "Now switch the USB switch back to the PC.\r\n" +
                "The app will capture the 'back' snapshot automatically in 10 seconds.";

            Log("Guided wizard: switch the USB switch back to the PC now. Capturing 'back' snapshot in 10 seconds.");
            _wizardTimer.Start();
            return;
        }

        if (_guidedWizardPhase == 2)
        {
            _snapshotBack = DeviceEnumerator.GetPresentDevices().Where(IsLikelyUsbish).ToList();
            _guidedWizardPhase = 0;

            Log($"Guided wizard: captured 'back' USB snapshot: {_snapshotBack.Count} devices.");
            UpdateWizardSummary();

            if (_trackedFamilies.Count > 0)
            {
                _switchState = SwitchState.OnPc;
                Log($"Guided wizard completed. Learned {_trackedFamilies.Count} tracked families.");
            }
            else
            {
                Log("Guided wizard completed, but no tracked families were learned.");
            }
        }
    }

    private void OnUsbTopologyChanged()
    {
        if (!_autoSwitchCheck.Checked)
            return;

        if (_guidedWizardPhase != 0)
            return;

        if (_isInCooldown)
            return;

        if (_trackedFamilies.Count == 0)
            return;

        _usbDebounceTimer.Stop();
        _usbDebounceTimer.Start();
    }

    private void UsbDebounceTimer_Tick(object? sender, EventArgs e)
    {
        _usbDebounceTimer.Stop();
        EvaluateAutoSwitch();
    }

    private void EvaluateAutoSwitch()
    {
        var currentFamilies = DeviceEnumerator.GetPresentDevices()
            .Where(IsLikelyUsbish)
            .Select(d => FamilyKey(d.InstanceId))
            .Where(f => _trackedFamilies.Contains(f))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        int trackedCount = _trackedFamilies.Count;
        int presentCount = currentFamilies.Count;
        int missingCount = trackedCount - presentCount;

        int awayThreshold = Math.Max(2, (int)Math.Ceiling(trackedCount * 0.35));
        int backThreshold = Math.Max(2, (int)Math.Ceiling(trackedCount * 0.25));

        Log($"Auto-switch eval: tracked={trackedCount}, present={presentCount}, missing={missingCount}, state={_switchState}");

        if ((_switchState == SwitchState.Unknown || _switchState == SwitchState.OnPc) && missingCount >= awayThreshold)
        {
            Log("Auto-switch: inferred USB switched away from PC.");
            AutoSwitchMonitors(toLaptop: true);
            _switchState = SwitchState.Away;
            EnterCooldown();
            return;
        }

        if (_switchState == SwitchState.Away && presentCount >= backThreshold)
        {
            Log("Auto-switch: inferred USB switched back to PC.");
            AutoSwitchMonitors(toLaptop: false);
            _switchState = SwitchState.OnPc;
            EnterCooldown();
        }
    }

    private void EnterCooldown()
    {
        _isInCooldown = true;
        _cooldownTimer.Stop();
        _cooldownTimer.Start();
        Log("Auto-switch cooldown started.");
    }

    private void CooldownTimer_Tick(object? sender, EventArgs e)
    {
        _cooldownTimer.Stop();
        _isInCooldown = false;
        Log("Auto-switch cooldown ended.");
    }

    private void AutoSwitchMonitors(bool toLaptop)
    {
        uint inputValue = toLaptop ? _config.LaptopInputValue : _config.PcInputValue;
        string label = toLaptop ? "Laptop" : "PC";

        RefreshMonitors();

        var selectedIds = _config.SelectedMonitorIds.ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var monitor in _monitors.Where(m =>
		m.IsAccessible &&
		(selectedIds.Count == 0 || selectedIds.Contains(GetMonitorSelectionKey(m)))))
        {
            bool ok = MonitorController.TrySetInput(monitor, inputValue);
            Log(ok
                ? $"Auto-switched '{monitor.BestName}' to {label} ({inputValue})."
                : $"Failed to auto-switch '{monitor.BestName}' to {label} ({inputValue}).");
        }
    }

    private void ReleaseMonitorHandles()
    {
        foreach (var monitor in _monitors)
        {
            if (monitor.PhysicalHandle != IntPtr.Zero)
            {
                NativeMethods.DestroyPhysicalMonitor(monitor.PhysicalHandle);
            }
        }

        _monitors.Clear();
    }

		private void RefreshMonitors()
		{
			_updatingMonitorChecks = true;
			try
			{
				ReleaseMonitorHandles();

				_monitors = MonitorController.GetMonitors();
				_monitorList.BeginUpdate();
				_monitorList.Items.Clear();

				var selectedIds = (_config.SelectedMonitorIds ?? new List<string>())
					.ToHashSet(StringComparer.OrdinalIgnoreCase);

				foreach (var m in _monitors)
				{
					var key = GetMonitorSelectionKey(m);

					bool isChecked = selectedIds.Count == 0
						? m.IsAccessible
						: selectedIds.Contains(key);

					var item = new ListViewItem(new[]
					{
						m.BestName,
						m.DeviceId,
						m.PhysicalDescription,
						m.IsAccessible ? "Yes" : "No"
					});

					item.Tag = m;
					_monitorList.Items.Add(item);
					item.Checked = isChecked;
				}

				EnsureMonitorColumns();
			}
			finally
			{
				_monitorList.EndUpdate();
				_updatingMonitorChecks = false;
			}

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

        if (toLaptop)
            _config.LaptopInputValue = value;
        else
            _config.PcInputValue = value;

        SaveConfig();

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

        _trackedFamilies = proposed.ToHashSet(StringComparer.OrdinalIgnoreCase);
        _config.TrackedFamilies = proposed;
        SaveConfig();

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

        sb.AppendLine();
        sb.AppendLine($"Tracked family count: {_trackedFamilies.Count}");

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
    _monitorList.Columns.Add("Input switch probe", 120);
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
	public string BusReportedDescription { get; init; } = "";
    public string MonitorDeviceName { get; init; } = "";
    public string PhysicalDescription { get; init; } = "";
    public string DeviceId { get; init; } = "";
    public string DeviceString { get; init; } = "";
    public string EdidName { get; init; } = "";
    public IntPtr HMonitor { get; init; }
    public IntPtr PhysicalHandle { get; init; }
    public bool HasPhysicalHandle { get; init; }
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

        var parts = deviceId.Split('\\');
        if (parts.Length < 2)
            return string.Empty;

        var model = parts[1];

        if (model.StartsWith("AUS", StringComparison.OrdinalIgnoreCase))
            return $"ASUS {model}";

        if (model.StartsWith("SAM", StringComparison.OrdinalIgnoreCase))
            return $"Samsung {model}";

        return model;
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
    private static bool TryGetDevicePropertyString(
    IntPtr infoSet,
    ref NativeMethods.SP_DEVINFO_DATA devInfo,
    NativeMethods.DEVPROPKEY propertyKey,
    out string value)
{
    value = string.Empty;

    NativeMethods.SetupDiGetDevicePropertyW(
        infoSet,
        ref devInfo,
        ref propertyKey,
        out _,
        IntPtr.Zero,
        0,
        out var requiredSize,
        0);

    if (requiredSize == 0)
        return false;

    var buffer = Marshal.AllocHGlobal((int)requiredSize);
    try
    {
        if (!NativeMethods.SetupDiGetDevicePropertyW(
            infoSet,
            ref devInfo,
            ref propertyKey,
            out _,
            buffer,
            requiredSize,
            out _,
            0))
            return false;

        value = Marshal.PtrToStringUni(buffer) ?? string.Empty;
        return !string.IsNullOrWhiteSpace(value);
    }
    finally
    {
        Marshal.FreeHGlobal(buffer);
    }
}

	public static string? TryGetBusReportedDescription(string deviceId)
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

            if (TryGetDevicePropertyString(
                infoSet,
                ref devInfo,
                NativeMethods.DEVPKEY_Device_BusReportedDeviceDesc,
                out var value))
            {
                return string.IsNullOrWhiteSpace(value) ? null : value;
            }

            return null;
        }
    }
    finally
    {
        NativeMethods.SetupDiDestroyDeviceInfoList(infoSet);
    }

    return null;
}
	
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
			var busReportedDescription = MonitorRegistryReader.TryGetBusReportedDescription(deviceId) ?? string.Empty;

            if (NativeMethods.GetNumberOfPhysicalMonitorsFromHMONITOR(hMonitor, out uint count) && count > 0)
            {
                var arr = new NativeMethods.PHYSICAL_MONITOR[count];
                if (NativeMethods.GetPhysicalMonitorsFromHMONITOR(hMonitor, count, arr))
                {
                    foreach (var pm in arr)
                    {
                        bool hasHandle = pm.hPhysicalMonitor != IntPtr.Zero;
                        bool supportsInputSwitch = hasHandle && TryProbeInputSwitch(pm.hPhysicalMonitor);

						result.Add(new MonitorInfo
						{
							DisplayName = mi.szDevice,
							MonitorDeviceName = mi.szDevice,
							PhysicalDescription = pm.szPhysicalMonitorDescription,
							DeviceId = deviceId,
							DeviceString = deviceString,
							EdidName = edidName,
							BusReportedDescription = busReportedDescription,
							HMonitor = hMonitor,
							PhysicalHandle = pm.hPhysicalMonitor,
							HasPhysicalHandle = hasHandle,
							IsAccessible = supportsInputSwitch,
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
        if (monitor.PhysicalHandle == IntPtr.Zero)
            return false;

        return NativeMethods.SetVCPFeature(monitor.PhysicalHandle, 0x60, inputValue);
    }

    private static bool TryProbeInputSwitch(IntPtr physicalHandle)
    {
        return NativeMethods.GetVCPFeatureAndVCPFeatureReply(
            physicalHandle,
            0x60,
            out _,
            out _,
            out _);
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

	public static readonly DEVPROPKEY DEVPKEY_Device_BusReportedDeviceDesc =
    new(new Guid("540b947e-8b40-45bc-a8a2-6a0b894cbda2"), 4);
	
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
    public static extern bool DestroyPhysicalMonitor(IntPtr hMonitor);

    [DllImport("dxva2.dll", SetLastError = true)]
    public static extern bool GetVCPFeatureAndVCPFeatureReply(
        IntPtr hMonitor,
        byte bVCPCode,
        out uint pvct,
        out uint pdwCurrentValue,
        out uint pdwMaximumValue);

    [DllImport("dxva2.dll", SetLastError = true)]
    public static extern bool SetVCPFeature(
        IntPtr hMonitor,
        byte bVCPCode,
        uint dwNewValue);
}