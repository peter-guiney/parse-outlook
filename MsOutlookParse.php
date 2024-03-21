
/**
 * Parse Outlook's Thread-Index header.
 *
 * @author peter.guiney@pentagoncomputers.com
 */
abstract class MsOutlookParse {
    public static function decodeThreadIndex(string $header): ?array
    {
        if (!$header) return null;

        try {
            // Decode the base64 string
            $decodedBytes = base64_decode($header);
            if (!$decodedBytes) return null;

            // Get this as hex
            $hex = bin2hex($decodedBytes);

            // This is the microsoft-documented way (which appears to only be the case for newer versions of outlook)

            // Skip the first reserved byte, then extract the next 5 bytes for FILETIME
            // Pad with three zero bytes on the right to make it 8 bytes long
            $fileTimeBytes = substr($decodedBytes, 1, 5) . "\x00\x00\x00";
            $asHex = bin2hex($fileTimeBytes);

            // Interpret these 8 bytes as a 64-bit unsigned integer (little endian) for the FILETIME
            //$fileTime = unpack('P', $fileTimeBytes)[1];
            $fileTime = hexdec($asHex);

            // Convert FILETIME to Unix timestamp
            $unixTimestamp = ($fileTime / 10000000) - 11644473600;

            // Check if our timestamp is sane
            if ($unixTimestamp < 0 || $unixTimestamp > (time() + 60*60*24*365*10)) {
                Log::warning('Thread index timestamp is out of range: ' . $unixTimestamp);

                // Instead go with the old undocumented way.
                // This is almost the same as above, but the 1st byte is NOT reserved, it's actually part of the
                // filetime (which hence has 6 bytes now, and only needs to be padded with 2 zero bytes).
                $fileTimeBytes = substr($decodedBytes, 0, 6) . "\x00\x00";
                $asHex = bin2hex($fileTimeBytes);
                $fileTime = hexdec($asHex);
                $unixTimestamp = ($fileTime / 10000000) - 11644473600;
            }

            // Extract the next 16 bytes for the GUID
            $guidBytes = substr($decodedBytes, 6, 16);
            // Convert the binary GUID into a hexadecimal string
            $guid = bin2hex($guidBytes);

            return ['id' => $guid, 'ts' => intval($unixTimestamp)];

        }
        catch (\Exception $e) {
            Log::warning('Error decoding thread index: ' . $e->getMessage());
            return null;
        }
    }
}
