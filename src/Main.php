<?php
include __DIR__ . '/../vendor/autoload.php';

const TYPE_LESSON = 1;
const TYPE_PAY_QUESTION = 2;

$filename = __DIR__ . '/../input/income.xlsx';
$dailyFilename = __DIR__ . '/../output/dailyResult.xlsx';
$weeklyFilename = __DIR__ . '/../output/weeklyResult.xlsx';

$reader = PHPExcel_IOFactory::createReaderForFile($filename);
$excel = $reader->load($filename);

$tableHeads = [];
$dataByDates = [];
$dataByWeeks = [];
$lessons = [];

/** @var PHPExcel_Worksheet_Row $row */
foreach ($excel->getActiveSheet()->getRowIterator() as $row) {

    if ($row->getRowIndex() === 1) {
        $tableHeads = parseHeads($row);
        continue;
    }

    $parsedRow = (parseRow($row, $tableHeads));
    if ($parsedRow) {
        $date = $parsedRow[0];
        $content = $parsedRow[2];
        $income = $parsedRow[4];
        $week = $parsedRow[5];

        if (!in_array($content, $lessons)) {
            $lessons[] = $content;
        }

        if (empty($dataByDates[$date])) {
            $dataByDates[$date] = [];
        }
        if (empty($dataByDates[$date][$content])) {
            $dataByDates[$date][$content] = 0;
        }
        $dataByDates[$date][$content] += $income;

        if (empty($dataByWeeks[$week])) {
            $dataByWeeks[$week] = [];
        }
        if (empty($dataByWeeks[$week][$content])) {
            $dataByWeeks[$week][$content] = 0;
        }
        $dataByWeeks[$week][$content] += $income;
    }
}

$dataByDates = array_reverse($dataByDates);
$dataByWeeks = array_reverse($dataByWeeks);

writeOutput($dataByDates, $lessons, $dailyFilename);
writeOutput($dataByWeeks, $lessons, $weeklyFilename);


function parseHeads(PHPExcel_Worksheet_Row $row)
{
    $tableHeads = [];

    /** @var PHPExcel_Cell $cell */
    foreach ($row->getCellIterator() as $cell) {
        $tableHeads[$cell->getColumn()] = $cell->getValue();
    }

    print_r($tableHeads);

    return $tableHeads;
}

function parseRow(PHPExcel_Worksheet_Row $row, array $tableHeads)
{
    $date = null;
    $type = null;
    $content = null;
    $isShared = '否';
    $income = null;
    $week = null;

    /** @var PHPExcel_Cell $cell */
    foreach ($row->getCellIterator() as $cell) {
        $value = $cell->getValue();
        switch ($tableHeads[$cell->getColumn()]) {
            case '时间':
                $date = PHPExcel_Style_NumberFormat::toFormattedString($value, PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDD);
                $week = date('W', strtotime($date));
                break;
            case '类型':
                if ($value == '讲堂报名费') {
                    $type = TYPE_LESSON;
                } else {
                    return null;
                }
                break;
            case '内容':
                if ($type == TYPE_LESSON) {
                    preg_match('/讲堂 "(.+?)" 的报名费(.*(推荐)?.*)/', $value, $matches);
                    $content = $matches[1];
                    if (!empty($matches[2])) {
                        $isShared = '是';
                    }
                }
                break;
            case '最终所得':
                $income = floatval($value);
                break;
            default:
                continue;
        }
    }

    return [$date, $type, $content, $isShared, $income, 'W' . $week];
}

function writeOutput($data, $lessons, $filename)
{
    $excel = new PHPExcel();
    $excel->setActiveSheetIndex(0)
        ->setCellValue('A1', '时间');
    $activeSheet = $excel->getActiveSheet();

    $row = 1;
    $columnMap = [];
    $columnIdx = 'A';
    foreach ($lessons as $lesson) {
        $nextColumnIdx = ++$columnIdx;
        $activeSheet->setCellValue($nextColumnIdx. 1, $lesson);
        $columnMap[$lesson] = $nextColumnIdx;
    }
    foreach ($data as $key => $record) {
        $activeSheet->setCellValue('A' . (++$row), $key);

        foreach ($record as $lesson => $income)  {
            foreach ($columnMap as $lessonKey => $column) {
                if ($lessonKey == $lesson) {
                    $activeSheet->setCellValue($column . $row, $income);
                }
            }
        }
    }

    $outputExcel = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
    $outputExcel->save($filename);
}

