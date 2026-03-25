using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using WinForms = System.Windows.Forms;

namespace AltSpaceApp
{
    public partial class MainWindow : Window
    {
        private const int HOTKEY_ID = 9000;
        private const uint MOD_ALT = 0x0001;
        private const uint VK_SPACE = 0x20;

        private ObservableCollection<AppInfo> apps = new ObservableCollection<AppInfo>();
        private ObservableCollection<AppInfo> filteredApps = new ObservableCollection<AppInfo>();

        private IntPtr windowHandle;
        private WinForms.NotifyIcon notifyIcon;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public MainWindow()
        {
            InitializeComponent();

            ResultsList.ItemsSource = filteredApps;
            ResultsList.KeyDown += ResultsList_KeyDown;
            ResultsList.MouseDoubleClick += ResultsList_MouseDoubleClick;
            SearchBox.TextChanged += SearchBox_TextChanged;
            SearchBox.KeyDown += SearchBox_KeyDown;

            LoadApps();

            InitializeTrayIcon();

            // 首次启动时隐藏，Alt+Space唤醒
            Hide();
        }

        private void ResultsList_MouseDoubleClick(object? sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (ResultsList.SelectedItem is AppInfo selected)
            {
                LaunchApp(selected);
            }
        }

        private void InitializeTrayIcon()
        {
            notifyIcon = new WinForms.NotifyIcon
            {
                Icon = System.Drawing.SystemIcons.Application,
                Text = "AltSpace Search",
                Visible = true
            };

            var menu = new WinForms.ContextMenuStrip();
            menu.Items.Add("Show", null, (s, e) => ShowWindow());
            menu.Items.Add("Exit", null, (s, e) => Application.Current.Shutdown());
            notifyIcon.ContextMenuStrip = menu;

            notifyIcon.DoubleClick += (s, e) => ShowWindow();
        }

        private void ShowWindow()
        {
            if (!IsVisible) Show();
            WindowState = WindowState.Normal;
            Activate();
            Topmost = true;
            SearchBox.Focus();
        }

        private void LoadApps()
        {
            apps.Clear();
            LoadAppsFromDirectory(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu) + "\\Programs");
            LoadAppsFromDirectory(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) + "\\Programs");
        }

        private void LoadAppsFromDirectory(string path)
        {
            if (!Directory.Exists(path)) return;

            try
            {
                foreach (string file in Directory.GetFiles(path, "*.lnk", SearchOption.AllDirectories))
                {
                    string appName = Path.GetFileNameWithoutExtension(file);
                    if (!string.IsNullOrWhiteSpace(appName) && apps.All(x => x.Name != appName))
                    {
                        apps.Add(new AppInfo { Name = appName, Path = file });
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                // 忽略访问失败的文件夹
            }
            catch (Exception)
            {
                // 其他错误忽略
            }
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string query = SearchBox.Text.Trim().ToLower();
            var results = apps.Where(app => app.Name.ToLower().Contains(query)).ToList();
            filteredApps.Clear();
            foreach (var item in results) filteredApps.Add(item);
            if (filteredApps.Any())
            {
                ResultsList.SelectedIndex = 0;
            }
        }

        private void SearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down)
            {
                if (filteredApps.Any())
                {
                    ResultsList.SelectedIndex = 0;
                    ResultsList.Focus();
                }
                return;
            }

            if (e.Key == Key.Enter)
            {
                // move focus to list (enter selection mode)
                if (filteredApps.Any())
                {
                    ResultsList.SelectedIndex = ResultsList.SelectedIndex >= 0 ? ResultsList.SelectedIndex : 0;
                    ResultsList.Focus();
                }
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Escape)
            {
                Hide();
            }
        }

        private void ResultsList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && ResultsList.SelectedItem is AppInfo selectedApp)
            {
                LaunchApp(selectedApp);
            }
            else if (e.Key == Key.Escape)
            {
                Hide();
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            // Handle Escape to hide the window when focused
            if (e.Key == Key.Escape)
            {
                Hide();
            }
        }

        private void LaunchApp(AppInfo app)
        {
            if (app == null || string.IsNullOrEmpty(app.Path)) return;

            try
            {
                Process.Start(new ProcessStartInfo(app.Path) { UseShellExecute = true });
                Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to launch {app.Name}: {ex.Message}", "Launch Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var source = (HwndSource)PresentationSource.FromVisual(this);
            windowHandle = source.Handle;
            source.AddHook(WndProc);

            bool hotkeyRegistered = RegisterHotKey(windowHandle, HOTKEY_ID, MOD_ALT, VK_SPACE);
            if (!hotkeyRegistered)
            {
                MessageBox.Show("Alt+Space 热键注册失败，请以管理员身份运行或检查是否已被占用。", "Hotkey Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_HOTKEY = 0x0312;
            if (msg == WM_HOTKEY && wParam.ToInt32() == HOTKEY_ID)
            {
                // Toggle visibility on hotkey
                ToggleWindow();
                handled = true;
            }
            return IntPtr.Zero;
        }

        private void ToggleWindow()
        {
            if (IsVisible && WindowState == WindowState.Normal)
            {
                Hide();
            }
            else
            {
                ShowWindow();
            }
        }

        protected override void OnDeactivated(EventArgs e)
        {
            base.OnDeactivated(e);
            Hide();
        }

        protected override void OnClosed(EventArgs e)
        {
            UnregisterHotKey(windowHandle, HOTKEY_ID);
            notifyIcon?.Dispose();
            base.OnClosed(e);
        }
    }

    public class AppInfo
    {
        public string Name { get; set; }
        public string Path { get; set; }
    }
}