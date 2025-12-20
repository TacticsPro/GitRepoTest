using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Office_Tools_Lite.Task_Helper;
public static partial class Get_UUID
{
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint GetSystemFirmwareTable(uint FirmwareTableProviderSignature, uint FirmwareTableID, nint pFirmwareTableBuffer, uint BufferSize);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern bool GetVolumeInformation(
        string lpRootPathName,
        StringBuilder lpVolumeNameBuffer,
        int nVolumeNameSize,
        out uint lpVolumeSerialNumber,
        out uint lpMaximumComponentLength,
        out uint lpFileSystemFlags,
        StringBuilder lpFileSystemNameBuffer,
        int nFileSystemNameSize);

    public static string GetMachineIdentifier()
    {
        try
        {
            const uint RSMB = 0x52534D42; // 'RSMB'
            var size = GetSystemFirmwareTable(RSMB, 0, nint.Zero, 0);
            if (size == 0)
                return GetDriveSerialOrFallback();

            var buffer = Marshal.AllocHGlobal((int)size);
            try
            {
                var read = GetSystemFirmwareTable(RSMB, 0, buffer, size);
                if (read == 0)
                    return GetDriveSerialOrFallback();

                var raw = new byte[size];
                Marshal.Copy(buffer, raw, 0, raw.Length);

                var offset = 8;
                while (offset + 4 < raw.Length)
                {
                    var type = raw[offset];
                    var length = raw[offset + 1];
                    if (type == 1 && length >= 0x19)
                    {
                        var uuidOffset = offset + 0x08;
                        var uuidBytes = new byte[16];
                        Array.Copy(raw, uuidOffset, uuidBytes, 0, 16);

                        var uuid = new Guid(uuidBytes).ToString("N").ToUpperInvariant();
                        if (uuid == "00000000000000000000000000000000" || uuid == "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
                            return GetDriveSerialOrFallback();

                        return uuid;
                    }

                    offset += length;
                    while (offset < raw.Length - 1 && (raw[offset] != 0 || raw[offset + 1] != 0))
                        offset++;
                    offset += 2;
                }
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to retrieve SMBIOS UUID.", ex);
        }

        return GetDriveSerialOrFallback();
    }

    private static string GetDriveSerialOrFallback()
    {
        try
        {
            var driveLetter = Path.GetPathRoot(Environment.SystemDirectory); // Usually "C:\\"
            uint serialNumber;
            GetVolumeInformation(driveLetter, null, 0, out serialNumber, out _, out _, null, 0);
            return serialNumber.ToString("X"); // e.g., "D3F5A8C4"
        }
        catch (Exception ex)
        {
            return Guid.NewGuid().ToString("N").ToUpperInvariant(); // Last resort
            throw new InvalidOperationException("Failed to retrieve drive serial number.", ex);
        }
    }
}
