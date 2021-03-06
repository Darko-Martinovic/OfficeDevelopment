﻿using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
// ReSharper disable RedundantAssignment

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

            var h = DetermineHasher((HashType) cmbHasher.SelectedItem);

            var fs = new FileStream(filePathNormalized, FileMode.Open, FileAccess.Read);

            var sha = new SHA256Managed();
            var byteHash = sha.ComputeHash(fs);
            fs.Position = 0;
            var resultHash = h.ComputeHash(fs);

            var sb = new StringBuilder();
            sb.AppendLine("File name - " + fd.FileName);
            sb.AppendLine(cmbHasher.SelectedItem.ToString()); // + ComputeHash(fs, HashType.Sha256));
            sb.AppendLine(Convert.ToBase64String(resultHash, 0, resultHash.Length));
            sb.AppendLine("Sha256");
            sb.AppendLine(Convert.ToBase64String(byteHash, 0, byteHash.Length));
            fs.Close();

            var fileResult = Path.Combine(Path.GetDirectoryName(filePath) ?? throw new InvalidOperationException(),"Hasher.txt");

            using (var outfile = new StreamWriter(fileResult, append: false, encoding: Encoding.UTF8))
            {
                outfile.Write(sb.ToString());
            }

            Process.Start("notepad.exe", fileResult);
            // ReSharper disable once RedundantAssignment
            sb = null;
            fs = null;
            h = null;
            sha = null;
            byteHash = null;


        }

        private enum HashType
        {
            Sha256,
            Sha384,
            Sha512,
            Md5,
            Sha1
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
            }

            return ha;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmbHasher.DataSource = Enum.GetValues(typeof(HashType));
            cmbHasher.SelectedItem = HashType.Sha256;
        }
    }
}
