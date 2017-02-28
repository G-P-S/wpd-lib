/*
 * Created by SharpDevelop.
 * User: jaran
 * Date: 05.11.2012
 * Time: 19:12
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using PortableDeviceApiLib;
using System.IO;
using System.Linq;
using WindowsPortableDevicesLib;

namespace WindowsPortableDevicesLib.Domain
{

    public class WindowsPortableDevice
    {
        private PortableDeviceClass device;

        public WindowsPortableDevice(string deviceID)
        {
            this.DeviceID = deviceID;
        }

        public string DeviceID { get; set; }
        public string FriendlyName { get; set; }
        public string Manufacturer { get; set; }
        public string Model { get; set; }
        public DevicePropertyKeys.WPD_DEVICE_TYPES DeviceType { get; set; }

        void ValidateConnection()
        {
            if (device == null)
                throw new InvalidOperationException("Not connected to device.");
        }

        public void Connect()
        {
            if (this.device != null)
                return;

            var clientInfo = (IPortableDeviceValues)new PortableDeviceTypesLib.PortableDeviceValuesClass();
            device = new PortableDeviceClass();
            device.Open(DeviceID, clientInfo);

            GetPropertiesFromDevice();

        }

        private void GetPropertiesFromDevice()
        {
            // Retrieve the properties of the device
            IPortableDeviceContent content;
            IPortableDeviceProperties properties;
            device.Content(out content);
            content.Properties(out properties);

            // Retrieve the values for the properties
            IPortableDeviceValues propertyValues;
            properties.GetValues("DEVICE", null, out propertyValues);

            string aString;
            propertyValues.GetStringValue(ref DevicePropertyKeys.WPD_DEVICE_FRIENDLY_NAME, out aString);
            this.FriendlyName = aString;
            propertyValues.GetStringValue(ref DevicePropertyKeys.WPD_DEVICE_MANUFACTURER, out aString);
            this.Manufacturer = aString;
            propertyValues.GetStringValue(ref DevicePropertyKeys.WPD_DEVICE_MODEL, out aString);
            this.Model = aString;

            int deviceType;
            propertyValues.GetSignedIntegerValue(ref DevicePropertyKeys.WPD_DEVICE_TYPE, out deviceType);
            this.DeviceType = (DevicePropertyKeys.WPD_DEVICE_TYPES)deviceType;
        }

        public void Disconnect()
        {
            if (!(this.device == null))
                return;
            
            this.device.Close();
            this.device = null;
        }

        public PortableDeviceFolder GetContentsOf(PortableDeviceFolder parent = null)
        {
            ValidateConnection();

            var root = new PortableDeviceFolder("DEVICE", "DEVICE", null);
            if (parent != null)
            {
                root = parent;
            }

            IPortableDeviceContent content;
            device.Content(out content);
            EnumerateContentsOf(ref content, root);

            return root;
        }

        public PortableDeviceFolder GetContentsOfRecursive(PortableDeviceFolder parent = null)
        {
            ValidateConnection();
  
            var root = new PortableDeviceFolder("DEVICE", "DEVICE", null);
            if (parent != null)
            {
                root = parent;
            }

            IPortableDeviceContent content;
            device.Content(out content);
            EnumerateContentsOfRecursive(ref content, root);

            return root;
        }

        //Interprets the string path as a forward-slash separated absolute path
        //and returns the PortableDeviceFolder object corresponding to the location
        //listed.
        //Returns null for path = "/". This way, calling one of the GetContents
        //functions above will make logical sense.
        //If the path is not found, this will throw an exception.
        public PortableDeviceFolder GetFolderFromPath(string path)
        {
            ValidateConnection();

            //Remove any leading and trailing slashes.
            while (true)
            {
                if (path[0] == '/')
                    path = path.Substring(1);
                else break;
            }
            while (true)
            {
                if (path[path.Length - 1] == '/')
                    path = path.Substring(0, path.Length - 1);
                else break;
            }

            if (path.Length == 0)
                return null;

            string[] directories = path.Split('/');
            string pathSoFar = "/";

            try
            {
                PortableDeviceFolder contents = this.GetContentsOfRecursive();
                PortableDeviceFolder thisLevel = contents;
                foreach (string folderToFind in directories)
                {
                    if (thisLevel.Files.Count == 0)
                        goto notFound;

                    bool folderFound = false;
                    foreach (PortableDeviceObject f in thisLevel.Files)
                        if (f.Name == folderToFind && f.GetType() == typeof(PortableDeviceFolder))
                        {
                            thisLevel = (PortableDeviceFolder)f;
                            folderFound = true;
                            break;
                        }

                    if (!folderFound)
                        goto notFound;

                    pathSoFar += folderToFind + "/";
                }

                return thisLevel;
            }
            catch (Exception exc)
            {
                throw new Exception(
                    "Exception during search for directory \"" + path + "\"\n" +
                    "pathSoFar = " + pathSoFar + "\nException message:\n" +
                    exc.Message);
            }

        notFound:
            throw new DirectoryNotFoundException("Failed to find the MTP path \"" + path + "\" at level \"" + pathSoFar + "\".");
        }

        public void GetFile(PortableDeviceFile file, string saveToPath)
        {
            ValidateConnection();

            IPortableDeviceContent content;
            device.Content(out content);

            IPortableDeviceResources resources;
            content.Transfer(out resources);

            PortableDeviceApiLib.IStream wpdStream = null;
            uint optimalTransferSize = 0;

            resources.GetStream(file.Id, ref DevicePropertyKeys.WPD_RESOURCE_DEFAULT, 0, ref optimalTransferSize, out wpdStream);

            System.Runtime.InteropServices.ComTypes.IStream sourceStream =
                (System.Runtime.InteropServices.ComTypes.IStream)wpdStream;

            var filename = Path.GetFileName(file.Name);

            try
            {
                using(FileStream targetStream = new FileStream(Path.Combine(saveToPath, filename), FileMode.Create, FileAccess.Write))
                {
                    unsafe
                    {
                        var buffer = new byte[optimalTransferSize];
                        int bytesRead;
                        do
                        {
                            sourceStream.Read(buffer, (int)optimalTransferSize, new IntPtr(&bytesRead));
                            targetStream.Write(buffer, 0, bytesRead);
                        } while (bytesRead > 0);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wpdStream);
            }
        }

        public PortableDeviceFile PutFile(string localFilePath, PortableDeviceFolder parent)
        {
            ValidateConnection();
  
            IPortableDeviceContent content;
            device.Content(out content);

            IPortableDeviceValues values = GetRequiredPropertiesForContentType(localFilePath, parent.Id);

            PortableDeviceApiLib.IStream tempStream;
            uint optimalTransferSizeBytes = 0;
            content.CreateObjectWithPropertiesAndData(values, out tempStream, ref optimalTransferSizeBytes, null);

            System.Runtime.InteropServices.ComTypes.IStream targetStream =
                (System.Runtime.InteropServices.ComTypes.IStream)tempStream;

            try
            {
                using (FileStream sourceStream = new FileStream(localFilePath, FileMode.Open, FileAccess.Read))
                {
                    var buffer = new byte[optimalTransferSizeBytes];
                    int bytesRead;
                    do
                    {
                        bytesRead = sourceStream.Read(buffer, 0, (int)optimalTransferSizeBytes);
                        IntPtr pcbWritten = IntPtr.Zero;
                        targetStream.Write(buffer, bytesRead, pcbWritten);
                    } while (bytesRead > 0);
                }
                targetStream.Commit(0);

                //Re-enumerate the contents of the folder, and return the new file.
                GetContentsOf(parent);
                foreach (PortableDeviceFile f in parent.Files)
                    if (f.Name == Path.GetFileName(localFilePath))
                        return f;

                return null;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tempStream);
            }
        }

        private IPortableDeviceValues GetRequiredPropertiesForContentType(string fileName, string parentObjectId)
        {
            IPortableDeviceValues values =
                new PortableDeviceTypesLib.PortableDeviceValues() as IPortableDeviceValues;

            values.SetStringValue(ref DevicePropertyKeys.WPD_OBJECT_PARENT_ID, parentObjectId);

            FileInfo fileInfo = new FileInfo(fileName);
            values.SetUnsignedLargeIntegerValue(DevicePropertyKeys.WPD_OBJECT_SIZE, (ulong)fileInfo.Length);
            values.SetStringValue(DevicePropertyKeys.WPD_OBJECT_ORIGINAL_FILE_NAME, Path.GetFileName(fileName));
            values.SetStringValue(DevicePropertyKeys.WPD_OBJECT_NAME, Path.GetFileName(fileName));
            values.SetBoolValue(DevicePropertyKeys.WPD_OBJECT_CAN_DELETE, 1);
            values.SetGuidValue(DevicePropertyKeys.WPD_OBJECT_FORMAT, DeviceGUIDS.WPD_CONTENT_TYPE_UNSPECIFIED);

            return values;
        }

        public override string ToString()
        {
            bool shouldDisconnect = false;
            if (device == null)
            {

                Connect();
                shouldDisconnect = true;
            }
            string toString = FriendlyName;

            if (shouldDisconnect)
            {

                Disconnect();
            }

            return toString;
        }

        public PortableDeviceObject GetObject(PortableDeviceFolder parent, string objectID)
        {
            IPortableDeviceContent content;
            device.Content(out content);

            // Get the properties of the object
            IPortableDeviceProperties properties;
            content.Properties(out properties);

            string parentID;

            if (parent == null)
                parentID = DeviceID;
            else
                parentID = parent.Id;

            // Enumerate the items contained by the current object
            IEnumPortableDeviceObjectIDs objectIds;
            content.EnumObjects(0, parentID, null, out objectIds);

            uint fetched = 0;
            do
            {
                string objectId;

                objectIds.Next(1, out objectId, ref fetched);
                if (fetched > 0)
                {
                    if (objectId.Equals(objectID))
                    {
                        PortableDeviceObject deviceObject = WrapObject(properties, objectId, parent);
                        if (deviceObject.GetType() == typeof(PortableDeviceFolder))
                        {
                            EnumerateContentsOf(ref content, (PortableDeviceFolder)deviceObject);

                        }
                        return deviceObject;
                    }
                }
            } while (fetched > 0);

            return null;
        }

        public PortableDeviceFolder CreateFolder(PortableDeviceFolder parent, string folderName)
        {
            ValidateConnection();

            IPortableDeviceContent content;
            device.Content(out content);

            // Get the properties of the object
            IPortableDeviceProperties properties;
            content.Properties(out properties);

            string parentId;

            if (parent == null)
                parentId = DeviceID;
            else
                parentId = parent.Id;

            IPortableDeviceValues createFolderValues = new PortableDeviceTypesLib.PortableDeviceValues() as IPortableDeviceValues;
            createFolderValues.SetStringValue(DevicePropertyKeys.WPD_OBJECT_PARENT_ID, parentId);
            createFolderValues.SetStringValue(DevicePropertyKeys.WPD_OBJECT_NAME, folderName);
            createFolderValues.SetStringValue(DevicePropertyKeys.WPD_OBJECT_ORIGINAL_FILE_NAME, folderName);
            createFolderValues.SetGuidValue(DevicePropertyKeys.WPD_OBJECT_CONTENT_TYPE, DeviceGUIDS.WPD_CONTENT_TYPE_FOLDER);
            createFolderValues.SetGuidValue(DevicePropertyKeys.WPD_OBJECT_FORMAT, DeviceGUIDS.WPD_OBJECT_FORMAT_PROPERTIES_ONLY);

            string newFolderId = "";
            content.CreateObjectWithPropertiesOnly(createFolderValues, ref newFolderId);

            //Re-enumerate the contents of the folder, and return the new file.
            GetContentsOf(parent);
            foreach (PortableDeviceFolder f in parent.Files)
                if (f.Name == folderName)
                    return f;

            return null;
        }

        public void DeleteFile(PortableDeviceFile file)
        {
            ValidateConnection();

            PortableDeviceFolder parent = (PortableDeviceFolder)file.Parent;

            IPortableDeviceValues values =
                new PortableDeviceTypesLib.PortableDeviceValues() as IPortableDeviceValues;

            values.SetStringValue(ref DevicePropertyKeys.WPD_OBJECT_ID, file.Id);

            tag_inner_PROPVARIANT propVariant = new tag_inner_PROPVARIANT();
            values.GetValue(ref DevicePropertyKeys.WPD_OBJECT_ID, out propVariant);

            IPortableDeviceContent content;
            device.Content(out content);

            IPortableDevicePropVariantCollection objectIds = new PortableDeviceTypesLib.PortableDevicePropVariantCollection() as IPortableDevicePropVariantCollection;
            objectIds.Add(propVariant);

            content.Delete(0, objectIds, null);

            GetContentsOf(parent);
        }

        private static void EnumerateContentsOf(ref IPortableDeviceContent content, PortableDeviceFolder deviceObject)
        {

            // Get the properties of the object
            IPortableDeviceProperties properties;
            content.Properties(out properties);

            // Enumerate the items contained by the current object
            IEnumPortableDeviceObjectIDs objectIds;
            content.EnumObjects(0, deviceObject.Id, null, out objectIds);

            deviceObject.Files.Clear();
            
            uint fetched = 0;
            do
            {
                string objectId;

                objectIds.Next(1, out objectId, ref fetched);
                if (fetched > 0)
                {
                    var currentObject = WrapObject(properties, objectId, deviceObject);

                    deviceObject.Files.Add(currentObject);

                }
            } while (fetched > 0);
        }

        private static void EnumerateContentsOfRecursive(ref IPortableDeviceContent content, PortableDeviceFolder deviceObject)
        {
            int fileCounter = 0;
            uint totalFetched = 0;
            // Get the properties of the object
            IPortableDeviceProperties properties;
            content.Properties(out properties);

            // Enumerate the items contained by the current object
            IEnumPortableDeviceObjectIDs objectIds;
            deviceObject.Files.Clear();
            content.EnumObjects(0, deviceObject.Id, null, out objectIds);

            uint fetched = 0;
            do
            {
                //Buggy MTP devices will cause issues during this recursive object search.
                //So we choose (for now) to swallow any errors.
                try
                {
                    string objectId;

                    objectIds.Next(1, out objectId, ref fetched);
                    totalFetched += fetched;
                    if (fetched > 0)
                    {
                        var currentObject = WrapObject(properties, objectId, deviceObject);

                        deviceObject.Files.Add(currentObject);

                        if (currentObject is PortableDeviceFolder)
                        {
                            Console.WriteLine("[wpd-lib] moving into: " + currentObject.GetFullPath());
                            EnumerateContentsOfRecursive(ref content, (PortableDeviceFolder)currentObject);
                        }
                        else
                        {
                            fileCounter++;
                            Console.WriteLine("[wpd-lib] " + fileCounter.ToString("0000") + " found a file: " + currentObject.GetFullPath());
                        }
                    }
                    else
                    {
                        Console.WriteLine("[wpd-lib] fetched zero objects in " + deviceObject.GetFullPath());
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("[wpd-lib] exception in EnumerateContentsOfRecursive: " + e.Message);
                }
            } while (fetched > 0);
        }

        private static PortableDeviceObject WrapObject(IPortableDeviceProperties properties, string objectId, PortableDeviceObject parent)
        {
            IPortableDeviceKeyCollection keys;
            properties.GetSupportedProperties(objectId, out keys);

            IPortableDeviceValues values;
            properties.GetValues(objectId, keys, out values);

            string name;
            values.GetStringValue(DevicePropertyKeys.WPD_OBJECT_NAME, out name);

            Guid contentType;
            values.GetGuidValue(DevicePropertyKeys.WPD_OBJECT_CONTENT_TYPE, out contentType);

            string uniqueID;
            values.GetStringValue(DevicePropertyKeys.WPD_OBJECT_PERSISTENT_UNIQUE_ID, out uniqueID);

            PortableDeviceObject deviceObject = null;

            if (contentType == DeviceGUIDS.WPD_CONTENT_TYPE_FOLDER || contentType == DeviceGUIDS.WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT)
            {
                deviceObject = new PortableDeviceFolder(objectId, name, parent);
            }
            else
            {
                deviceObject = new PortableDeviceFile(objectId, name, parent);
            }

            deviceObject.PersistentId = uniqueID;

            return deviceObject;
        }
    }
}
