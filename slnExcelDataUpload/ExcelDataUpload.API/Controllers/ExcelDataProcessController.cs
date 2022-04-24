using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ExcelDataReader;
using ExcelDataUpload.API.Helper;
using Microsoft.Extensions.Configuration;

namespace ExcelDataUpload.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelDataProcessController : ControllerBase
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
       
        private readonly DataHandling objHelperDataAccess;

        public ExcelDataProcessController(IWebHostEnvironment webHostEnvironment, IConfiguration configuration)
        {
            _webHostEnvironment = webHostEnvironment ?? throw new ArgumentNullException(nameof(webHostEnvironment));
            objHelperDataAccess = new DataHandling(configuration);
        }

        [Route("UploadExcelData")]
        [HttpPost]
        public IActionResult ProcessExcelFile(List<IFormFile> files)
        {
            try
            {
                if (files.Count == 0)
                    return BadRequest();

                string message = "";

                DataSet dsexcelRecordsAll = new DataSet();

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                foreach (var file in files)
                {
                    IFormFile Inputfile = null;
                    Stream FileStream = null;
                    IExcelDataReader reader = null;
                    DataSet dsexcelRecords = new DataSet();

                    Inputfile = file;
                    
                    FileStream = Inputfile.OpenReadStream();

                    if (Inputfile != null && FileStream != null)
                    {
                        if (Inputfile.FileName.EndsWith(".xls"))
                        { reader = ExcelReaderFactory.CreateBinaryReader(FileStream); }
                        else if (Inputfile.FileName.EndsWith(".xlsx"))
                        { reader = ExcelReaderFactory.CreateOpenXmlReader(FileStream); }
                        else { message = "The file format is not supported."; }

                        dsexcelRecords = reader.AsDataSet();
                        reader.Close();

                        if (dsexcelRecords != null && dsexcelRecords.Tables.Count > 0)
                        {
                            objHelperDataAccess.ProcessDataAdotoSQL(dsexcelRecords, "sampleTable1", null, false);
                            //DataTable dtExcelRecords = dsexcelRecords.Tables[0];

                            //Write Function to Convert copy data from bulk Copy.
                            //dsexcelRecordsAll.Tables.Add(dsexcelRecords.Tables[0]);
                        }
                        else { message = "Selected file is empty."; }
                    }
                    else { message = "Invalid File."; }
                }

                return Ok(message);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
