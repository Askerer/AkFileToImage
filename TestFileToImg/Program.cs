using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FileToImage;

namespace TestFileToImg
{
	class Program
	{
		static void Main(string[] args)
		{
			FileToImage.FileTransform a = new FileTransform();
			//a.ExcelToJPG("");
			a.WordToJPG("");

		}
	}
}
