<?php

namespace Artisan\Console\Commands;

use Illuminate\Console\Command;

class AttendanceExcel extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'excel:attendance {file : Excel文件路径} {--month= : 统计日期开始的月份} {--cycle= : 考勤周期开始日期}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = '按人统计月考勤结果';

    /**
     * 上下班时间
     * @var string
     */
    protected $daily_stime = '10:01:00';
    protected $daily_etime = '19:00:00';

    /**
     * 月份
     * @var [type]
     */
    protected $month;

    /**
     * 超过多久算加班
     * @var integer
     */
    protected $overtime_minutes = 60;

    /**
     * 考勤周期开始日期
     * @var integer
     */
    protected $cycle = 26;

    /**
     * 调休上班
     * @var [type]
     */
    protected $special_date = [
        // 周末上班
        'work' => [
            '2018-12-29', // 元旦调休上班 周六
            '2019-02-02', // 春节调休上班 周六
            '2019-02-03', // 春节调休上班 周日
            '2019-09-29', // 国庆调休上班 周日
            '2019-10-12', // 国庆调休上班 周六
        ],
        // 工作日休息
        'off'  => [
            '2018-12-31',  // 元旦调休 周一
            '2019-01-01',  // 元旦 周二
            '2019-02-04',  // 除夕 周一
            '2019-02-05',  // 春节 周二
            '2019-02-06',
            '2019-02-07',
            '2019-02-08',
            '2019-04-05',  // 清明节 周五
            '2019-05-01',  // 劳动节 周三
            '2019-06-07',  // 端午节 周五
            '2019-09-13',  // 中秋节 周五
            '2019-10-01',  // 国庆节 周二
            '2019-10-02',
            '2019-10-03',
            '2019-10-04',
            '2019-10-07',
        ],
    ];

    /**
     * reader path
     * @var [type]
     */
    protected $reader_path = __DIR__.'/../../script/excelreader.js';

    /**
     * writer path
     * @var [type]
     */
    protected $writer_path = __DIR__.'/../../script/excelwriter.js';

    /**
     * data path
     * @var [type]
     */
    protected $data_path   = __DIR__.'/../../data/attendance/';

    /**
     * 周几MAPPER
     */
    const DAY_MAPPER = [
        1 => '星期一',
        2 => '星期二',
        3 => '星期三',
        4 => '星期四',
        5 => '星期五',
        6 => '星期六',
        7 => '星期日',
    ];

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        ini_set('date.timezone','Etc/GMT');
        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return mixed
     */
    public function handle()
    {
        $file = $this->argument('file');
        $this->month = sprintf("%02d", $this->option('month')?:(date('m') - 1));
        if ($this->option('cycle')) {
            $this->cycle = $this->option('cycle');
        }
        $data = $this->getdata($file);
        $path = $this->data_path.date('Y-'.sprintf("%02d", intval($this->month) + 1)).'/';
        foreach ($data as $sheet) {
            $statdata = [
                'overtime' => [['日期', '星期', '加班开始时间', '加班结束时间', '加班小时']],
                'late'     => [['日期', '星期', '上班打卡时间', '下班打卡时间', '调休开始时间','调休结束时间', '调休小时','备注']],
            ];
            unset($sheet['data'][0]);
            $sheet['data'] = $this->formatData($sheet['data']);
            foreach ($sheet['data'] as $row) {
                $stat = $this->getDayStat($row);
                $statdata['overtime'][] = $stat[0];
                $statdata['late'][] = $stat[1];
            }
            $statdata = [
                [
                    'name' => '加班',
                    'data' => $statdata['overtime'],
                ],
                [
                    'name' => '调休Or忘打卡',
                    'data' => $statdata['late'],
                ],
            ];
            $statdata[0]['data'][] = [
                '姓名',$sheet['name'], '', '合计', '=SUM(E2:E'.count($statdata[0]['data']).')'
            ];
            $statdata[1]['data'][] = [
                '姓名',$sheet['name'], '', '', '','合计','=SUM(G2:G'.count($statdata[1]['data']).')'
            ];

            $this->writedata($statdata, $path, (intval($this->month) + 1).'月考勤记录_'.$sheet['name']);
        }
        echo "统计已生成 ".$path.PHP_EOL;
        system('open '.$path);
    }

    /**
     * 格式化数据
     * @param  [type] $data [description]
     * @return [type]       [description]
     */
    public function formatData($data){
        $month = $this->month;
        $time = strtotime(date('Y-'.($month).'-'.$this->cycle));
        while ($time < strtotime(date('Y-'.($month + 1).'-'.$this->cycle))) {
            $e_data[date('Y-m-d', $time)]['date'] = date('Y-m-d', $time);
            $e_data[date('Y-m-d', $time)]['time'] = [];
            $time += 24 * 3600;
        }
        foreach ($data as $value) {
            if (is_float($value[3])) {
                $value[3] = date('Y/m/d H:i:s', ($value[3] - 25569) * 24*60*60);
            }
            if (isset($e_data[date('Y-m-d', strtotime($value[3]))])) {
                $e_data[date('Y-m-d', strtotime($value[3]))]['time'][] = $value[3];
            }
        }
        return $e_data;
    }

    /**
     * 单天统计
     * @param  [type] $row [description]
     * @return [type]      [description]
     */
    public function getDayStat($row){
        return [
            $this->getOvertime($row),
            $this->getLate($row),
        ];
    }

    /**
     * 加班
     * @param  [type] $row [description]
     * @return [type]      [description]
     */
    public function getOvertime($row){
        // 有打卡记录 且超时
        if ($row['time']) {
            // 周末 调休休息 || 工作日 调休休息
            if ((in_array(date('N', strtotime($row['date'])), [6, 7])  && !in_array(date('Y-m-d', strtotime($row['date'])), $this->special_date['work']))
            || (in_array(date('N', strtotime($row['date'])), [1, 2, 3, 4, 5])  && in_array(date('Y-m-d', strtotime($row['date'])), $this->special_date['off']))) {
                if (count($row['time']) >= 2) {
                    // 整天加班
                    $hours = $this->calOverTime(end($row['time']), $row['time'][0]);
                    $data = [
                        $row['date'],
                        self::DAY_MAPPER[date('N', strtotime($row['date']))],
                        $hours ? date('H:i', strtotime($row['time'][0])): '',
                        $hours ? date('H:i', strtotime(end($row['time']))) : '',
                        $hours ? : '',
                    ];
                } else {
                    // 周末只打一次不算加班
                    $data = [
                        $row['date'],
                        self::DAY_MAPPER[date('N', strtotime($row['date']))],
                        '',
                        '',
                        ''
                    ];
                }
                
            } else {
                // 计算加班
                $hours = $this->calOverTime(end($row['time']));
                $data = [
                    $row['date'],
                    self::DAY_MAPPER[date('N', strtotime($row['date']))],
                    $hours ? date('H:i', strtotime('19:00')) : '',
                    $hours ? date('H:i', strtotime(end($row['time']))) : '',
                    $hours ? : '',
                ];
            }
        } else {
            // 未加班
            $data = [
                $row['date'],
                self::DAY_MAPPER[date('N', strtotime($row['date']))],
                '',
                '',
                ''
            ];
        }
        return $data;
    }

    /**
     * 计算加班时间
     * @param  [type]  $to   [description]
     * @param  integer $from [description]
     * @return [type]        [description]
     */
    public function calOverTime($to, $from = 0){
        $from = $from?:date('Y-m-d 19:00:00',strtotime($to));
        $sec = strtotime($to) - strtotime($from);
        if ($sec > $this->overtime_minutes * 60) {
            // 超过规定时间算加班
            $h_hours = floor($sec / 1800);
            return $h_hours / 2;
        } else {
            return 0;
        }
    }

    /**
     * 迟到早退未打卡
     * @param  [type] $row [description]
     * @return [type]      [description]
     */
    public function getLate($row){
        // 周末 休息 || 工作日休息 加班没有异常
        if ((in_array(date('N', strtotime($row['date'])), [6, 7])  && !in_array(date('Y-m-d', strtotime($row['date'])), $this->special_date['work']))
            || (in_array(date('N', strtotime($row['date'])), [1, 2, 3, 4, 5])  && in_array(date('Y-m-d', strtotime($row['date'])), $this->special_date['off']))) {
            // 非工作日加班 正常
            return [
                $row['date'],
                self::DAY_MAPPER[date('N', strtotime($row['date']))],
                '',
                '',
                ''
            ];
        }

        if (count($row['time']) >= 2) {
            // 迟到或早退
            if ( strtotime($row['time'][0])   > strtotime($row['date'].' '.$this->daily_stime) || 
                 strtotime(end($row['time'])) < strtotime($row['date'].' '.$this->daily_etime)) {
                // 第一次打卡已经下班了
                if ( strtotime($row['time'][0])   > strtotime($row['date'].' '.$this->daily_etime)){
                    return [
                        $row['date'],
                        self::DAY_MAPPER[date('N', strtotime($row['date']))],
                        '未打卡',
                        end($row['time']),
                        '',
                    ];
                // 最后一次打卡还没上班
                } elseif ( strtotime(end($row['time'])) < strtotime($row['date'].' '.$this->daily_stime) ){
                    return [
                        $row['date'],
                        self::DAY_MAPPER[date('N', strtotime($row['date']))],
                        $row['time'][0],
                        '未打卡',
                        '',
                    ];
                } else {
                    return [
                        $row['date'],
                        self::DAY_MAPPER[date('N', strtotime($row['date']))],
                        $row['time'][0],
                        end($row['time']),
                        '',
                    ];
                }
                
            } else {
                // 正常
                return [
                    $row['date'],
                    self::DAY_MAPPER[date('N', strtotime($row['date']))],
                    '',
                    '',
                    ''
                ];
            }
        } elseif(count($row['time']) == 1) {
            // 忘打卡
            if (strtotime($row['time'][0]) <= strtotime($row['date'].' '.$this->daily_stime )) {
                return [
                    $row['date'],
                    self::DAY_MAPPER[date('N', strtotime($row['date']))],
                    $row['time'][0],
                    '未打卡',
                    '',
                ];
            } elseif (strtotime($row['time'][0]) >= strtotime($row['date'].' '.$this->daily_etime )) { // 早退
                return [
                    $row['date'],
                    self::DAY_MAPPER[date('N', strtotime($row['date']))],
                    '未打卡',
                    $row['time'][0],
                    '',
                ];
            } else {
                return [
                    $row['date'],
                    self::DAY_MAPPER[date('N', strtotime($row['date']))],
                    $row['time'][0],
                    '未打卡',
                    '',
                ];
            }
        } else { // 请假全天
            return [
                $row['date'],
                self::DAY_MAPPER[date('N', strtotime($row['date']))],
                '未打卡',
                '未打卡',
                '',
            ];
        }
        
    }

    /**
     * 计算迟到时间
     * @param  [type] $from [description]
     * @param  [type] $to   [description]
     * @return [type]       [description]
     */
    public function calLateHours($from, $to){
        $sec = strtotime($to) - strtotime($from);
        $h_hours = ceil($sec / 1800);
        return $h_hours / 2;
    }

    public function getdata($file = ''){
        if (!$file) {
            $file = $this->data_path.'origin.xlsx ';
        }
        if (!file_exists($file)) {
            die('文件未找到，请使用绝对路径');
        }
        $cmd = 'node '.$this->reader_path.' '.$file.' '.$this->data_path.'origin.json';
        system($cmd);
        $json = file_get_contents($this->data_path.'origin.json');
        unlink($this->data_path.'origin.json');
        return json_decode($json, true);
    }

    public function writedata($data, $path, $name){
        $json_path = $path.$name.'.json';
        $xlsx_path = $path.$name.'.xlsx';
        if (!is_dir($path)) {
            mkdir($path);
        }
        file_put_contents($json_path, json_encode($data, true));
        $cmd = 'node '.$this->writer_path.' '.$json_path.' '.$xlsx_path;
        system($cmd);
        unlink($json_path);
        return true;
    }
}
