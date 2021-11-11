<?php

namespace Manager\Excel;

use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class EManager{

    public function __construct()
    {

    }

    public static function Hello()
    {
        echo "hello composer";
    }

    /**
     * @desc 【组件消息传递】
     *       返回数据或结束进程
     * @param $status -1 返回错误数据信息 / 2 返回数据异常,终止
     * @param array $data
     * @param string $msg
     * @return array
     */
    public function importReturnMsg($status, $data = [], $msg = '')
    {
        $return = [
            'status' => $status,
            'payload' => [
                'data' => empty($data) ? [] : $data,
                'msg' => $msg,
            ],
        ];
        if ($status == '-1') return $return;
        die(json_encode($return, 320));
    }

    /**
     * @param string $template 模板路径
     * @param array $data 模板替换数据
     * @param string $output 输出 空-打印 路径-保存,返回文件名
     * @return string|void
     */
    public function templateExcelExport($template, $data, $output='')
    {
        // 验证数据格式
        $data = array_values($data);    # 去除原有数据的键,重排序为索引数组
        if (!file_exists($template)) $this->importReturnMsg('2', [], '模板【' . $template . '】不存在');
        try {
            $templateType = IOFactory::identify($template);     # 获取读取模板文件类型对应读取器类型
            $reader = IOFactory::createReader($templateType);   # 生成对应读取器
            $reader->setLoadAllSheets();                        # 设置读取全部sheet页
            $spreadsheets = $reader->load($template);           # 读取模板数据
        } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
            $this->importReturnMsg('2', [], json_encode($e->getMessage(), 320));
        }
        Calculation::getInstance($spreadsheets)->clearCalculationCache();   # 刷新计算公式
        // 循环检测数据包格式, 是否有复制sheet单页需求
        try {
            $this->copySheetIfNeeded($spreadsheets, $data);
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            $this->importReturnMsg('2', [], json_encode($e->getMessage(), 320));
        }
        // 开始每一sheet页处理,替换数据
        $workSheetNames = $spreadsheets->getSheetNames();
        foreach ($workSheetNames as $sheetIndex => $workSheetName) {
            try {
                $currentWorkSheet = $spreadsheets->getSheet($sheetIndex);  # 依照sheet索引加载单独sheet
            } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
                $this->importReturnMsg('2', [], json_encode($e->getMessage(), 320));
            }
            $currentSheetData = $currentWorkSheet->toArray();                 # 获取当前sheet中所有数据
            $currentMainData = $data[$sheetIndex]['main'] ?? [];
            $currentLoopData = $data[$sheetIndex]['loop'] ?? [];
            try {
                $this->replaceMainData($currentWorkSheet, $currentSheetData, $currentMainData);     # 循环替换主表数据
                $this->replaceLoopData($currentWorkSheet, $currentLoopData);                        # 循环替换循环数据
            } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
                $this->importReturnMsg('2', [], json_encode($e->getMessage(), 320));
            }
        }
        // 设置第一页为当前页
        try {
            $spreadsheets->setActiveSheetIndex(0);
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            $this->importReturnMsg('2', [], json_encode($e->getMessage(), 320));
        }
        // 命名
        $file_path_info = pathinfo($template);
        $file_name = $file_path_info['filename'] . date('YmdHis', time()) . mt_rand() . '.' . strtolower($templateType);
        // 下载/保存本地
        if (empty($output)) {
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename=' . $file_name);
            header('Cache-Control: max-age=0');
            ob_clean();
            flush();
            try {
                $writer = IOFactory::createWriter($spreadsheets, $templateType);
                $writer->save('php://output');
            } catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
                $this->importReturnMsg('2', [], json_encode($e->getMessage(), 320));
            }
            #__ 释放内存
            $spreadsheets->disconnectWorksheets();
            ob_end_flush();
            return;
        } else {
            $savePath = $output . '/' . $file_name;
            try {
                $writer = IOFactory::createWriter($spreadsheets, $templateType);
                $writer->save($savePath);
            } catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
                $this->importReturnMsg('2', [], json_encode($e->getMessage(), 320));
            }
            #__ 释放内存
            $spreadsheets->disconnectWorksheets();
            return $savePath;
        }
    }

    /**
     * @desc 如果需要复制,复制sheet页
     * @param Spreadsheet $spreadsheets
     * @param $data
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function copySheetIfNeeded(Spreadsheet $spreadsheets, $data)
    {
        foreach ($data as $dataIndex => $datum) {   # 循环数据包, 是否有需要copy的数据
            if (empty($currentCopyData)) continue;
            $currentCopyData = $datum['copy'];
            $copiedIndex = !empty($currentCopyData['index']) ? $currentCopyData['index']-1 : 0;   # 复制对象索引,默认0
            $copiedTitle = $currentCopyData['title'];       # 复制后sheet标题
            $copiedWorkSheet = $spreadsheets->getSheet($copiedIndex);   # 依照sheet索引加载单独sheet
            $newWorksheet = clone $copiedWorkSheet;                     # 克隆新的sheet

            // 如果没有指定title,复制原复制对象的title,加 (数字)
            if (empty($copiedTitle)){
                $copiedWorksheetName = $copiedWorkSheet->getTitle();
                $nowExistNames = $spreadsheets->getSheetNames();
                $extCount = 1;
                do {
                    $copiedTitle = "{$copiedWorksheetName}({$extCount})";
                    $extCount ++;
                } while ( in_array($copiedTitle, $nowExistNames) );
            }

            $newWorksheet->setTitle($copiedTitle);
            $spreadsheets->addSheet($newWorksheet, $dataIndex);
        }
    }

    /**
     * @desc 替换单次替换主值(图片+文本)
     * @param Worksheet $worksheet  当前sheet页对象
     * @param array $sheetData      当前sheet页数据数组
     * @param array $data           替换数组
     *                              # 文本类替换数据 替换格式为: $xmmc$ 的数据
     *                              # 图片类替换数据 替换格式为: $@qrcode$ 或 $@qrcode[50:60:50:50:50]$ 的数据
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function replaceMainData(Worksheet $worksheet, $sheetData, $data)
    {
        $emptyLine = 0;
        foreach ($sheetData as $lineIndex => $sheetLineDatum) {
            if ($emptyLine > 10) break;
            $isEmptyLine = true;
            foreach ($sheetLineDatum as $cellIndex => $cellDatum) {
                if ( $cellDatum == null ) continue;
                $isEmptyLine = false;
                if (strpos($cellDatum,'$') === false || preg_match('/^\$.+\$$/', $cellDatum) === 0) continue;     # 没有$符号$,跳过
                # 获取模板标识符
                $cellSign = str_replace('$', '', $cellDatum);
                # 处理单元格数据
                $this->manageSingleCellData($worksheet, $lineIndex, $cellIndex, $cellSign, $data);
            }
            $emptyLine = $isEmptyLine == true ? $emptyLine + 1 : 0; # 若当前行为空行, 空行数+1, 否则,空行数清空
        }
    }

    /**
     * @desc 寻找 #前缀.字段名# 格式所在行及 前缀名
     * @param Worksheet $worksheet
     * @return array|bool
     */
    protected function findLoopStart(Worksheet $worksheet)
    {
        $sheetData = $worksheet->toArray();     # 当前excel数组
        $emptyLine = 0;
        foreach ($sheetData as $lineIndex => $sheetLineDatum) {
            if ($emptyLine > 10) break;  # 累积空行超过10行,停止
            $isEmptyLine = true;
            foreach ($sheetLineDatum as $cellIndex => $cellDatum) {
                if ( $cellDatum == null ) continue;
                $isEmptyLine = false;
                if (strpos($cellDatum, '#') !==false && preg_match('/^#.+#$/', $cellDatum) !== 0){
                    $sign = str_replace('#','',$cellDatum);
                    $signDate = explode('.',$sign);
                    if ( isset($signDate[1]) && (empty($signDate[0]) || empty($signDate[1]) || isset($signDate[2]) ) )
                        $this->importReturnMsg('2', [], json_encode(["标识符[ $sign ]不是有效的循环标识符,字段中可能存在符号[.]"], 320));
                    $signPre = $signDate[0];
                    if (count($signDate) == 1) $signPre = '';   # 若标识符中没有符号点, 代表没有前缀, 单表循环
                    return [ 'line' => $lineIndex,'pre' => $signPre ];
                }
            }
            $emptyLine = $isEmptyLine == true ? $emptyLine + 1 : 0; # 若当前行为空行, 空行数+1, 否则,空行数清空
        }
        return false;
    }

    /**
     * @desc 替换循环数据  #前缀.字段名# (带多子表替换循环)
     * @param Worksheet $worksheet
     * @param array $data 如果为单循环表,格式为[0=>[第一条数据],1=>[第二条数据],...], 如果为多循环表, 格式为 ['表1前缀名'=>[0=>[第一条数据],1=>[第二条数据],...],'表2前缀名'=>[...],...]
     *                                      ↑                             ↑                                  ↑                             ↑
     *                                      |------------------------------|                                   |-----------------------------|
     *                                                    |------------------→→→→→→→→→→→→→→→→→→→→→---------------|
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function replaceLoopData(Worksheet $worksheet, $data)
    {
        $startLine = $this->findLoopStart($worksheet);
        if ($startLine !== false){
            $rowPre     =  $startLine['pre'];   # 循环数据前缀
            $rowStart   = $startLine['line'];   # 循环起始行
            $sheetData = $worksheet->toArray();     # 当前excel数组
            $realRow = $rowStart + 1;
            $loopData = empty($rowPre) ? $data : ($data[$rowPre]??[]);
            $dataCount = count($loopData);
            if (!empty($loopData)) {
                $rowHeight = $worksheet->getRowDimension($realRow)->getRowHeight();
                $worksheet->insertNewRowBefore($realRow, $dataCount);   # 在起始插入行前插入新行
                for ($i = 0 ; $i < $dataCount ; $i ++) $worksheet->getRowDimension($realRow+$i)->setRowHeight($rowHeight); # 为新行设置行高
                // 使用旧表格行数据做对照循环赋值新行
                foreach ($sheetData[$rowStart] as $fieldIndex => $sheetDatum) {
                    $fieldValue = strpos($sheetDatum, '#') === false ? $sheetDatum : '';   # 如果不是模板标识符,使用原值,否则置空
                    $fieldName = str_replace(['#'.$rowPre.'.','#'],'',$sheetDatum);        # 去除 #前缀. 和 # 符号
                    foreach ($loopData as $dataIndex => $loopDatum) {
                        $this->manageSingleCellData($worksheet, $rowStart+$dataIndex, $fieldIndex, $fieldName, $loopDatum);
                        if (!empty($fieldValue)) $this->insertTextToWorksheet($worksheet, $rowStart+$dataIndex, $fieldIndex, $fieldValue);
                    }
                }
            }
            $worksheet->removeRow($realRow + $dataCount);
            $this->replaceLoopData($worksheet, $data);
        }
    }

    /**
     * @desc 处理单行数据中单个单元格的图片/文字填充
     * @param Worksheet $worksheet 当前sheet页对象
     * @param number $lineIndex 从0开始计数的单元格行数索引
     * @param number $cellIndex 从0开始计数的单元格列数索引
     * @param string $cellSign  单元格模板标识符
     * @param array $oneLineData 单行数据包
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function manageSingleCellData(Worksheet $worksheet, $lineIndex, $cellIndex, $cellSign, $oneLineData)
    {
        if (strpos($cellSign, '@') !== false){  # 处理图片(标识符中包含@符)
            preg_match('/\[(.*?)\]/',$cellSign,$match);     # 匹配第一个中括号内容
            if ( isset( $oneLineData[$cellSign] ) ) {  # 若有可用的值且文件存在,格式为 @tupian
                $imgPath = $oneLineData[$cellSign];
                $formatValue = [];
            }elseif(!empty($match)) {       # 无该格式值,且有 @tupian[10:10:5:5:20]
                $signValue = strstr($cellSign, '[', true);
                $imgPath = $oneLineData[$signValue] ?? '';
                $formatValue = strstr($cellSign, '[');                      # 获取[]内部数据
                $formatValue = str_replace(['[',']'], '', $formatValue);    # 去除括号
                $formatValue = explode(':',$formatValue);                   # 转化为数组格式
            }else{
                $imgPath = '';
                $formatValue = [];
            }

            $this->insertImageToWorksheet($worksheet, $lineIndex, $cellIndex, $imgPath, $formatValue);

        }else{      # 处理文本
            $textValue = isset($oneLineData[$cellSign]) ? $oneLineData[$cellSign] : '';
            if ($textValue === null) $textValue = '';
            $this->insertTextToWorksheet($worksheet, $lineIndex, $cellIndex, $textValue);
        }
    }

    /**
     * @desc 向excel对象中添加图片
     * @param Worksheet $worksheet      当前sheet页对象
     * @param number    $lineIndex      从0开始计数的单元格行数索引
     * @param number    $cellIndex      从0开始计数的单元格列数索引
     * @param string    $imgPath        图片的绝对路径 TODO 引入Image,保存线上图片到本地
     * @param array     $imgFormat      支持对图片的格式编辑,['宽','高','左偏移','上偏移','所在行高']
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function insertImageToWorksheet(Worksheet $worksheet, $lineIndex, $cellIndex, $imgPath = '', $imgFormat = [])
    {
        # 转义索引为真实索引
        $realRow = $lineIndex + 1;
        $col = $cellIndex + 1;
        $realCol = Coordinate::stringFromColumnIndex($col);
        # 如果带图片格式, 处理参数
        if (!empty($imgFormat)){
            $img_width = $imgFormat[0];
            $img_height = isset($imgFormat[1]) ? $imgFormat[1] : $imgFormat[0];     # 若仅一位,同长同宽
            $img_offset_x = isset($imgFormat[2]) ? $imgFormat[2] : 0;
            $img_offset_y = isset($imgFormat[3]) ? $imgFormat[3] : 0;
            if (isset($imgFormat[4])){
                $line_height = $imgFormat[4];
                $worksheet->getRowDimension($realRow)->setRowHeight($line_height);  # 设置行高
            }
        }
        # 依据参数决定添加图片的方式
        if (!empty($imgPath) && file_exists($imgPath)){   # 存在图片且文件为本地文件
            $objDrawing = new Drawing();
            $objDrawing->setPath($imgPath);
            if (!empty($imgFormat)){
                $objDrawing->setOffsetY($img_offset_y); //交给实施去慢慢调吧,
                $objDrawing->setOffsetX($img_offset_x);
                $objDrawing->setResizeProportional(false);
                $objDrawing->setWidthAndHeight($img_width, $img_height);
            }
            $objDrawing->setCoordinates($realCol . $realRow);
            $objDrawing->setWorksheet($worksheet);
        }elseif(!file_exists($imgPath) && strpos($imgPath, 'http') !== false && preg_match('/^http.*/', $imgPath) !== 0){   # 文件在本地不存在,且为http开头
            # 使用GD函数创建图片资源resource
            $gdImageString = file_get_contents($imgPath);
            $gdImage = imagecreatefromstring($gdImageString);
            $pathInfo = pathinfo($imgPath);
            $imageExtension = $pathInfo['extension'];
            if ($imageExtension == 'jpg') $imageExtension = 'jpeg';  # 适配渲染器和mime类型
            $imageName =  $pathInfo['filename'] . '.' . $pathInfo['extension'];
            # 添加内存图片到sheet页
            $drawing = new MemoryDrawing();
            $drawing->setName($imageName);
            $drawing->setDescription('In-Memory image');
            $drawing->setCoordinates($realCol . $realRow);
            $drawing->setImageResource($gdImage);
            // $drawing->setRenderingFunction(MemoryDrawing::RENDERING_JPEG);
            $drawing->setRenderingFunction('image'.$imageExtension);  # 根据文件扩展名设置渲染功能
            // $drawing->setMimeType(MemoryDrawing::MIMETYPE_DEFAULT);
            $drawing->setMimeType('image/'.$imageExtension);  # 根据文件扩展名设置mime类型
            if (!empty($imgFormat)){
                $drawing->setResizeProportional(false);     # 关闭比例保持
                $drawing->setWidthAndHeight($img_width, $img_height);
                $drawing->setOffsetX($img_offset_x);
                $drawing->setOffsetY($img_offset_y);
            }
            $drawing->setWorksheet($worksheet);
        }
        $worksheet->setCellValueByColumnAndRow($col, $realRow, ''); # 将图片原有标识符置空
    }

    /**
     * @desc 向当前sheet指定单元格添加文本数据(全文本格式)
     * @param Worksheet $worksheet  当前sheet页对象
     * @param number $lineIndex     从0开始计数的单元格行数索引
     * @param number $cellIndex     从0开始计数的单元格列数索引
     * @param string $textValue     需要写入的
     * @param string $format
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function insertTextToWorksheet(Worksheet $worksheet, $lineIndex, $cellIndex, $textValue='', $format='')
    {
        $row = $lineIndex + 1;
        $col = $cellIndex + 1;
        $realCol = Coordinate::stringFromColumnIndex($col);
        $cellRowCol = $realCol . $row;
        $worksheet->getStyle($cellRowCol)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_TEXT);
        $worksheet->getCell($cellRowCol)->setValueExplicit($textValue, DataType::TYPE_STRING);
    }

}