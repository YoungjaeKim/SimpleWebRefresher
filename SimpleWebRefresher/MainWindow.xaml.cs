using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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

namespace SimpleWebRefresher
{

	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private bool isUrlValue;
		private int refreshCount;
		public MainWindow()
		{
			InitializeComponent();

			// http://stackoverflow.com/a/6198700/361100
			WebBrowserMain.Navigated += new NavigatedEventHandler(wbMain_Navigated);

			var dispatcherTimer = new System.Windows.Threading.DispatcherTimer();

			dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
			int interval = 1;
			int.TryParse(TextboxTimer.Text.Trim(), out interval);
			dispatcherTimer.Interval = new TimeSpan(0, 0, interval);
			dispatcherTimer.Start();
		}


		private void dispatcherTimer_Tick(object sender, EventArgs e)
		{
			Uri url;
			try
			{
				url = new Uri(TextBoxAddress.Text.Trim());
			}
			catch (Exception)
			{
				TextBlockLog.Text = "주소 오류";
				refreshCount = 0;
				return;
			}
			WebBrowserMain.Navigate(url);

		}

		void wbMain_Navigated(object sender, NavigationEventArgs e)
		{
			SetSilent(WebBrowserMain, true); // make it silent

			TextBlockLog.Text = String.Format("{0}번 리프래시됨", ++refreshCount);
		}


		public static void SetSilent(WebBrowser browser, bool silent)
		{
			if (browser == null)
				throw new ArgumentNullException("browser");

			// get an IWebBrowser2 from the document
			IOleServiceProvider sp = browser.Document as IOleServiceProvider;
			if (sp != null)
			{
				Guid IID_IWebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");
				Guid IID_IWebBrowser2 = new Guid("D30C1661-CDAF-11d0-8A3E-00C04FC9E26E");

				object webBrowser;
				sp.QueryService(ref IID_IWebBrowserApp, ref IID_IWebBrowser2, out webBrowser);
				if (webBrowser != null)
				{
					webBrowser.GetType().InvokeMember("Silent", BindingFlags.Instance | BindingFlags.Public | BindingFlags.PutDispProperty, null, webBrowser, new object[] { silent });
				}
			}
		}


		[ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		private interface IOleServiceProvider
		{
			[PreserveSig]
			int QueryService([In] ref Guid guidService, [In] ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);
		}
	}
}
