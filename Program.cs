using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SetDesktopWallpaper
{
	public class Program
	{
		private readonly Random mRandom = new Random();

		// Create a list of file names from the root directory and its subdirectories
		private List<string> FetchFiles(string root)
		{
			var fileList = new List<string>(1);

			// Place the path of the root folder into the queue (FIFO)
			var pending = new Queue<string>(1);
			pending.Enqueue(root);

			while (pending.Count > 0)
			{
				// Get the path of the first folder from the queue
				string path = pending.Dequeue();
				string[] files = null;

				// Find all the files from this folder
				try
				{
					files = Directory.GetFiles(path);
				}
				catch { }

				// If the folder has at least one file, add the paths of all files in this folder
				if (files?.Length > 0)
					fileList.AddRange(files);

				// Place the paths of all subfolders into the queue
				try
				{
					foreach (string subdir in Directory.GetDirectories(path))
						pending.Enqueue(subdir);
				}
				catch { }
			}

			return fileList;
		}

		// Pick a random image file from the root directory and its subdirectories
		// and set this file as desktop wallpaper
		public bool FindAndSetWallpaper(string dir)
		{
			// Create a list of file names from the root directory and its subdirectories
			Console.Write($"Fetching files from {dir} ... ");
			List<string> files = FetchFiles(dir);

			if (files.Count == 0)
			{
				Console.WriteLine($"\nNo file exists in {dir}");
				return false;
			}

			Console.WriteLine();

			// Obtain the dimensions of the display device (e.g. monitor)
			const float tolerance = 489f / 2200f;

			Rectangle bounds = Screen.PrimaryScreen.Bounds;
			int screenHeight = bounds.Height;
			int screenWidth = bounds.Width;

			// Obtain the aspect ratio of the display device
			float screenRatio = (float)screenHeight / screenWidth;
			Image img;

			// Try to find the right image to be set as desktop background (in no more than 50 attempts)
			for (byte attempt = 1; attempt <= 50; attempt++)
			{
				int numFiles = files.Count;

				// If there is no file to choose, abort
				if (numFiles == 0)
					break;

				// Pick a random file without replacement
				int index = numFiles == 1 ? 0 : mRandom.Next(numFiles);
				string filename = files[index];
				files.RemoveAt(index);

				// Check whether the file is an eligible image file
				Console.Write($"\nTesting file: {filename}");

				try
				{
					img = Image.FromFile(filename);
				}
				catch
				{
					// If not, find another file
					Console.WriteLine($"\nInvalid image file: {filename}");
					continue;
				}

				// Obtain the dimensions of the image
				int height = img.Height;

				// If the image does not fit the screen, find another one
				if (height < screenHeight)
				{
					Console.WriteLine("\nDimensions not fit.");
					img.Dispose();
					continue;
				}

				int width = img.Width;
				img.Dispose();

				if (width < screenWidth)
				{
					Console.WriteLine("\nDimensions not fit.");
					continue;
				}

				// Test whether the aspect ratio of the image is close enough to fit the whole screen
				float aspectRatio = (float)height / width;
				float aspectToScreen = aspectRatio / screenRatio - 1f;

				// If not, find another image
				if (Math.Abs(aspectToScreen) > tolerance)
				{
					Console.WriteLine("\nDimensions not fit.");
					continue;
				}

				// Try to set wallpaper with the image as desktop background
				// If successful, terminate this program
				Console.Write($"\nSetting wallpaper: {filename}");

				if (SetDesktopWallpaper(filename))
				{
					Console.WriteLine("\nWallpaper set successfully.");
					return true;
				}

				Console.WriteLine($"\nUnable to set wallpaper: {filename}");
			}

			// Notify the user of failure
			Console.WriteLine("\nUnable to find an image file.");
			return false;
		}

		// Set wallpaper as desktop background using a given file name
		public static bool SetDesktopWallpaper(string filename)
		{
			// Use registry to set wallpaper in fill style
			RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true);
			key.SetValue(@"WallpaperStyle", "10");
			key.SetValue(@"TileWallpaper", "0");
			key.Close();

			// Set wallpaper as desktop background
			return SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, filename, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
		}

		[STAThread]
		public static void Main()
		{
			// Read the path of the folder previously chosen by user
			string myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			string rootFolderDir = Path.Combine(myDocuments, "SetDesktopWallpaper");
			Directory.CreateDirectory(rootFolderDir);

			string rootFolderFile = Path.Combine(rootFolderDir, "rootFolder.txt");
			string selectedPath = null;

			if (File.Exists(rootFolderFile))
				selectedPath = File.ReadAllText(rootFolderFile, Encoding.UTF8);

			// Ask for a root folder
			FolderBrowserDialog dialog = new FolderBrowserDialog
			{
				Description = "Please choose a root folder below:",
				ShowNewFolderButton = false
			};

			if (!string.IsNullOrEmpty(selectedPath))
				dialog.SelectedPath = selectedPath;

			// If cancelled, terminate the program
			if (dialog.ShowDialog() != DialogResult.OK)
			{
				dialog.Dispose();
				return;
			}

			selectedPath = dialog.SelectedPath;
			dialog.Dispose();

			// Find an image file from the folder selected by user
			// and set wallpaper as desktop background
			Program prog = new Program();
			bool success = prog.FindAndSetWallpaper(selectedPath);

			// If successful, save the path of the selected folder to file
			if (success)
				File.WriteAllText(rootFolderFile, selectedPath, Encoding.UTF8);

			// Prompt the user to exit the program (for debug purpose only)
			// Console.Write("\nPress any key to terminate ... ");
			// Console.ReadKey();
		}

		private const uint SPI_SETDESKWALLPAPER = 20;
		private const uint SPIF_UPDATEINIFILE = 0x01;
		private const uint SPIF_SENDWININICHANGE = 0x02;

		[DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]

		private static extern bool SystemParametersInfo(uint uiAction, uint uiParam,
			string pvParam, uint fWinIni);
	 }
}
