using System;
using System.Runtime.InteropServices;

namespace COMInteropWebExample
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void btnCreateFile_Click(object sender, EventArgs e)
        {
            // ProgID for FileSystemObject
            const string progId = "Scripting.FileSystemObject";
            Type fsoType = null;
            object fso = null;

            try
            {
                // Get the COM type for FileSystemObject
                fsoType = Type.GetTypeFromProgID(progId);
                if (fsoType == null)
                {
                    lblStatus.Text = "FileSystemObject is not available on this system.";
                    return;
                }

                // Create an instance of FileSystemObject
                fso = Activator.CreateInstance(fsoType);

                // Create a text file
                string filePath = Server.MapPath("~/App_Data/COMInteropExample.txt"); // Save in App_Data folder
                object file = fsoType.InvokeMember(
                    "CreateTextFile",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    fso,
                    new object[] { filePath, true } // File path and overwrite flag
                );

                // Write to the file
                file.GetType().InvokeMember(
                    "WriteLine",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    file,
                    new object[] { "Hello, this is written using COM interop with FileSystemObject." }
                );

                // Close the file
                file.GetType().InvokeMember(
                    "Close",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    file,
                    null
                );

                lblStatus.Text = $"File created successfully at: {filePath}";
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"Error: {ex.Message}";
            }
            finally
            {
                // Clean up
                if (fso != null)
                {
                    Marshal.ReleaseComObject(fso);
                }
            }
        }
    }
}
