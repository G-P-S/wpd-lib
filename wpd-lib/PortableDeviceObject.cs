/*
 * Created by SharpDevelop.
 * User: jaran
 * Date: 08.11.2012
 * Time: 18:41
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

namespace WindowsPortableDevicesLib.Domain
{
    /// <summary>
    /// Description of PortableDeviceObject.
    /// </summary>
    public class PortableDeviceObject
    {
                
        protected PortableDeviceObject(string id, string name, PortableDeviceObject parent)
        {
            this.Id = id;
            this.Name = name;
            this.Parent = parent;
        }

        public string PersistentId { get; set; }

        public string Id { get; private set; }

        public string Name { get; private set; }

        public PortableDeviceObject Parent { get; private set; }

        public string GetFullPath()
        {
            if (this.Parent == null)
                return "/";

            string path = "/" + this.Name;

            PortableDeviceObject o = Parent;
            while (o.Parent != null)
            {
                path = "/" + o.Name + path;
                o = o.Parent;
            }

            return path;
        }
    }
}
