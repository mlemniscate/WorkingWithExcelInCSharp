using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ExcelDemoAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;
        private readonly WeatherForecast[] weatherForecasts;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
            weatherForecasts = Enumerable.Range(1, 5).Select(index => new WeatherForecast
                {
                    Date = DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
                    TemperatureC = Random.Shared.Next(-20, 55),
                    Summary = Summaries[Random.Shared.Next(Summaries.Length)]
                })
                .ToArray();
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return weatherForecasts;
        }

        // [HttpGet(Name = "GetWeatherForecastExcel")]
        [HttpGet("excel", Name = "GetWeatherForecastExcel")]
        public async Task<IActionResult> GetExcel()
        {
            string workingDirectory = Environment.CurrentDirectory;
            // string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            var file = new FileInfo(workingDirectory + @"\Files\WeatherExcel.xlsx");
            await SaveExcelFile(weatherForecasts, file);
            return Ok();
        }

        private async Task SaveExcelFile(WeatherForecast[] weatherForecasts, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add("Report");

            var range = ws.Cells["A1"].LoadFromCollection(weatherForecasts, true);
            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }
    }
}