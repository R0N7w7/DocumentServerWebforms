using System;
using System.Security.Cryptography;
using System.Text;

namespace WebEditor
{
    public static class OnlyOfficeJwt
    {
        public static string Create(string jsonPayload, string secret)
        {
            if (string.IsNullOrWhiteSpace(jsonPayload)) throw new ArgumentException("Payload is required", nameof(jsonPayload));
            if (secret == null) secret = string.Empty;

            var headerJson = "{\"alg\":\"HS256\",\"typ\":\"JWT\"}";
            var header = Base64UrlEncode(Encoding.UTF8.GetBytes(headerJson));
            var payload = Base64UrlEncode(Encoding.UTF8.GetBytes(jsonPayload));
            var signingInput = header + "." + payload;
            var signatureBytes = HmacSha256(Encoding.UTF8.GetBytes(secret), Encoding.UTF8.GetBytes(signingInput));
            var signature = Base64UrlEncode(signatureBytes);
            return signingInput + "." + signature;
        }

        public static bool Validate(string token, string secret)
        {
            if (string.IsNullOrWhiteSpace(token)) return false;
            if (secret == null) secret = string.Empty;

            var parts = token.Split('.');
            if (parts.Length != 3) return false;

            var signingInput = parts[0] + "." + parts[1];
            var expectedSig = Base64UrlEncode(HmacSha256(Encoding.UTF8.GetBytes(secret), Encoding.UTF8.GetBytes(signingInput)));
            var actualSig = parts[2];
            return FixedTimeEquals(Encoding.ASCII.GetBytes(expectedSig), Encoding.ASCII.GetBytes(actualSig));
        }

        private static string Base64UrlEncode(byte[] input)
        {
            return Convert.ToBase64String(input)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        private static byte[] HmacSha256(byte[] key, byte[] data)
        {
            using (var h = new HMACSHA256(key))
            {
                return h.ComputeHash(data);
            }
        }

        private static bool FixedTimeEquals(byte[] a, byte[] b)
        {
            if (a == null || b == null || a.Length != b.Length) return false;
            var diff = 0;
            for (int i = 0; i < a.Length; i++) diff |= a[i] ^ b[i];
            return diff == 0;
        }
    }
}
