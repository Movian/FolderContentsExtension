using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.Linq;

namespace FolderContentsExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Directory)]
    public class FolderContentsExtension : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            // Enable the menu only when there are files in the selected folder
            return GetFilesInFolder().Count > 0;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            // Create the context menu
            var menu = new ContextMenuStrip();

            // Create the sub-menu
            var subMenu = new ToolStripMenuItem("Contents");

            // Retrieve the list of files in the folder you right-click on
            List<string> files = GetFilesInFolder();

            foreach (string file in files)
            {
                Icon fileIcon;
                Image UseIcon;
                try
                {


                    // Get the file icon and set it as the menu item's image
                    fileIcon = Shellicon.GetLargeIcon(file);
                    if(fileIcon == null)
                    {
                        fileIcon = Shellicon.GetSmallIcon(file);
                    }

                    // Check if the icon has a valid image
                    if (fileIcon?.Size.Width > 0 && fileIcon?.Size.Height > 0)
                    {
                        UseIcon = fileIcon.ToBitmap();
                    }
                    else
                    {
                        UseIcon = SystemIcons.Application.ToBitmap();
                    }

                    // Create a context menu item for each file
                    var fileMenuItem = new ToolStripMenuItem
                    {
                        Image = UseIcon,
                        Text = "     - " + Path.GetFileName(file)
                    };

                    fileMenuItem.Click += (sender, e) =>
                    {
                        // Handle the click event for the file menu item
                        // Implement the desired action when the user selects the menu item
                        try
                        {
                            Process.Start(file);
                        }
                        catch (Exception ex)
                        {
                            // Log the exception
                            LogError($"Error opening file: {file}", ex);
                        }
                    };

                    // Add the file menu item to the sub-menu
                    subMenu.DropDownItems.Add(fileMenuItem);
                }
                catch (Exception ex)
                {
                    // Log the exception
                    LogError($"Error creating menu item for file: {file}", ex);
                }
            }

            // Add the sub-menu to the main context menu
            menu.Items.Add(subMenu);

            return menu;
        }

        private void LogImageSize(Icon FileIcon)
        {
            string logFilePath = @"C:\Temp\FolderContentsExtension_ErrorLog.txt";
            string logMessage = $"{DateTime.Now}: Width: {FileIcon?.Size.Width}\nHeight: {FileIcon?.Size.Height}\n\n";
            File.AppendAllText(logFilePath, logMessage);
        }


        private List<string> GetFilesInFolder()
        {
            // Implement the logic to retrieve the list of files in the folder
            string folderPath = SelectedItemPaths.FirstOrDefault();
            return new List<string>(Directory.GetFiles(folderPath));
        }

        private void LogError(string message, Exception ex)
        {
            // Log the error message and exception to a file or console
            string logFilePath = @"C:\Temp\FolderContentsExtension_ErrorLog.txt";
            string logMessage = $"{DateTime.Now}: {message}\nException: {ex}\n\n";

            // Write the log message to a file
            File.AppendAllText(logFilePath, logMessage);
        }


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool DestroyIcon(IntPtr handle);


    }
}
