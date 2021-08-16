<?php
namespace App\Console\Commands;

require 'vendor/autoload.php';

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Log;
use Symfony\Component\Console\Output\ConsoleOutput;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


class GenerateReport extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'report:generate {cities}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Command description';

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        $output = new ConsoleOutput();
        $cities = explode(',', $this->argument('cities'));
        $row = 1;

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'City');
        $sheet->setCellValue('B1', 'Date');
        $sheet->setCellValue('C1', 'Average Humidity');
        $sheet->setCellValue('D1', 'Average Temperature');
        $sheet->setCellValue('E1', 'Chance of Rain');

        foreach ($cities as $city) {
            try {
                $curl = curl_init("http://api.weatherapi.com/v1/forecast.json?key=2e00c87a08354d84822111056211408&q={$city}&days=3&aqi=no&alerts=no");
                curl_setopt($curl, CURLOPT_HEADER, 0);
                curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);

                $result = curl_exec($curl);
                $data = json_decode($result);

                if (property_exists($data, 'error')) {
                    $output->writeln('One ore more cities cannot be found.');
                    exit;
                }

                foreach ($data->forecast->forecastday as $fd) {
                    $row+=1;
                    $sheet->setCellValue('A'.$row, $city);
                    $sheet->setCellValue('B'.$row, $fd->date);
                    $sheet->setCellValue('C'.$row, $fd->day->avghumidity);
                    $sheet->setCellValue('D'.$row, $fd->day->avgtemp_c.' ('.$fd->day->avgtemp_f.')');
                    $sheet->setCellValue('E'.$row, $fd->day->daily_chance_of_rain.'%');
                }
                $row+=1;

                curl_close($curl);
            } catch (Exception $ex) {
                return $ex->getMessage();
            }
        }

        $output->writeln('Report successfully generated.');
        $writer = new Xlsx($spreadsheet);
        $writer->save('weather.xlsx');
    }
}
