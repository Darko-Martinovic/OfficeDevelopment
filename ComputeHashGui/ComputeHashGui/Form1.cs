using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace ComputeHashGui
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void BtnFile_Click(object sender, EventArgs e)
        {

            var fd = new OpenFileDialog { Title = @"Choose a file" };

            if (fd.ShowDialog() != DialogResult.OK) return;

            txtFile.Text = fd.FileName;
            var filePath = fd.FileName;
            var filePathNormalized = Path.GetFullPath(filePath);
            var fs = new FileStream(filePathNormalized, FileMode.Open, FileAccess.Read);
            var sb = new StringBuilder();
            sb.AppendLine("File name - " + fd.FileName);
            sb.AppendLine(Environment.NewLine);
            sb.AppendLine("SHA1- " + ComputeHash(fs, HashType.Sha1));
            sb.AppendLine(Environment.NewLine);
            sb.AppendLine("SHA256- " + ComputeHash(fs, HashType.Sha256));
            sb.AppendLine(Environment.NewLine);
            sb.AppendLine("SHA384- " + ComputeHash(fs, HashType.Sha384));
            sb.AppendLine(Environment.NewLine);
            sb.AppendLine("SHA512- " + ComputeHash(fs, HashType.Sha512));
            sb.AppendLine(Environment.NewLine);
            sb.AppendLine("MD5- " + ComputeHash(fs, HashType.Md5));
            sb.AppendLine(Environment.NewLine);
            sb.AppendLine("Ripemd160- " + ComputeHash(fs, HashType.Ripemd160));
            fs.Close();
            string fileResult = Path.GetDirectoryName(fd.FileName) + "Hasher_" +
                                DateTime.Now.ToString("dd_MM_yyyy_hh_mm_ss") + ".txt";

            using (var outfile = new StreamWriter(fileResult, append: false, encoding: Encoding.UTF8))
            {
                outfile.Write(sb.ToString());
            }

            Process.Start("notepad.exe", fileResult);
            // ReSharper disable once RedundantAssignment
            sb = null;



        }

        private enum HashType
        {
            Sha1,
            Sha256,
            Sha384,
            Sha512,
            Md5,
            Ripemd160
        }

        private static HashAlgorithm DetermineHasher(HashType h)
        {
            HashAlgorithm ha = null;
            switch (h)
            {
                case HashType.Sha1:
                    ha = new SHA1CryptoServiceProvider();
                    break;
                case HashType.Sha256:
                    ha = new SHA256Managed();
                    break;
                case HashType.Sha384:
                    ha = new SHA384Managed();
                    break;
                case HashType.Sha512:
                    ha = new SHA512Managed();
                    break;
                case HashType.Md5:
                    ha = new MD5CryptoServiceProvider();
                    break;
                case HashType.Ripemd160:
                    ha = new RIPEMD160Managed();
                    break;
            }

            return ha;
        }

        private static string ComputeHash(Stream fs, HashType hs)
        {
            var ha = DetermineHasher(hs);
            var byteHash = ha.ComputeHash(fs);
            return Convert.ToBase64String(byteHash, 0, byteHash.Length);
        }

    }
}
