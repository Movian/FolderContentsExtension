using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

namespace FolderContentsExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Directory)]
    public class FolderContentsExtension : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            // Enable the menu only when there are files in the selected folder
            return GetFilesInFolder().Any();
        }

        protected override ContextMenuStrip CreateMenu()
        {
            // Create the main context menu
            var menu = new ContextMenuStrip();

            // Create the sub-menu named "Contents"
            var subMenu = new ToolStripMenuItem("Contents");

            // Set the maximum height of the sub-menu to display a scroll bar if needed
            subMenu.DropDown.MaximumSize = new System.Drawing.Size(400, 0); // Adjust the maximum width as needed

            // Retrieve the list of files in the folder you right-click on
            List<string> files = GetFilesInFolder();

            foreach (string file in files)
            {
                // Create a context menu item for each file
                var fileMenuItem = new ToolStripMenuItem(file);
                fileMenuItem.Click += (sender, e) =>
                {
                    // Handle the click event for the file menu item
                    // Open the file with its default behavior
                    Process.Start(file);
                };

                // Add the file menu item to the sub-menu
                subMenu.DropDownItems.Add(fileMenuItem);
            }

            // Add the sub-menu to the main context menu
            menu.Items.Add(subMenu);

            return menu;
        }

        private List<string> GetFilesInFolder()
        {
            // Implement the logic to retrieve the list of files in the folder
            string folderPath = SelectedItemPaths.FirstOrDefault();
            return new List<string>(Directory.GetFiles(folderPath));
        }
    }
}
