<?php
/**
 * MySQL数据库导出Excel模型
 */

// 服务器地址
$servername = '';
// 账户
$username = '';
// 密码
$password = '';
// 数据库
$db = '';
// Excel名称
$title = '';

// 创建连接
$conn = new mysqli($servername, $username, $password);

// 检测连接
if ($conn->connect_error) {
    die("连接失败: " . $conn->connect_error);
}

// 获取数据库内所有表名以及注释
$table_res = $conn->query("select table_name, table_comment from information_schema.TABLES where table_schema = '{$db}'");
$tables = [];

while($table = $table_res->fetch_assoc()){

    // 查询表索引
    $index_res = $conn->query("show index from {$db}.{$table['table_name']}");
    $indexs = '';
    while($index = $index_res->fetch_assoc()){
        $indexs .= $index['Column_name'] . ',';
    }
    $indexs = trim($indexs, ',');

    // 查询表信息
    $column_res = $conn->query("select COLUMN_NAME, IS_NULLABLE, DATA_TYPE, COLUMN_TYPE, EXTRA, COLUMN_COMMENT from information_schema.columns where table_schema = '{$db}' and TABLE_NAME = '{$table['table_name']}'");
    $columns = [];
    while($column = $column_res->fetch_assoc()){
        $columns[] = $column;
    }

    // 数据组装
    $tables[] = [
        'name' => $table['table_name'],
        'comment' => $table['table_comment'],
        'field' => $columns,
        'index' => $indexs
    ];
}

// 引入Excel
require "PHPExcel-1.8/Classes/PHPExcel.php";

$obpe = new \PHPExcel();
$obpe->setactivesheetindex();

// 设置Excel名称
$obpe->getActiveSheet()->setTitle($title);

// 边框
$styleThinBlackBorderOutline = array(
    'borders' => array(
        'allborders' => array( //设置全部边框
            'style' => \PHPExcel_Style_Border::BORDER_THIN //粗的是thick
        ),

    ),
);

$k = 0;

foreach ($tables as $key => $table)
{
    // 与上一个表格保持两行的距离
    $k = $k + 2;

    // 此表格第一行k
    $first_key = $k;

    // 表格第一行 表名和注释
    $obpe->getActiveSheet()->mergeCells('B' . $k . ':C' . $k);
    $obpe->getActiveSheet()->mergeCells('E' . $k . ':F' . $k);
    $obpe->getactivesheet()->setcellvalue('A' . $k, '表名');
    $obpe->getactivesheet()->setcellvalue('B' . $k, $table['name']);
    $obpe->getactivesheet()->setcellvalue('D' . $k, '注释');
    $obpe->getactivesheet()->setcellvalue('E' . $k, $table['comment']);
    $k++;

    // 表格第二行 主键
    $obpe->getActiveSheet()->mergeCells('B' . $k . ':F' . $k);
    $obpe->getactivesheet()->setcellvalue('A' . $k, '索引');
    $obpe->getactivesheet()->setcellvalue('B' . $k, $table['index']);
    $k++;

    // 表格第三行 字段标题
    $obpe->getactivesheet()->setcellvalue('A' . $k, '字段名');
    $obpe->getactivesheet()->setcellvalue('B' . $k, '数据类型');
    $obpe->getactivesheet()->setcellvalue('C' . $k, '为空');
    $obpe->getActiveSheet()->mergeCells('D' . $k . ':F' . $k);
    $obpe->getactivesheet()->setcellvalue('D' . $k, '注释');
    $k++;

    // 表格第四段 开始表字段输出
    foreach ($table['field'] as $field)
    {
        $obpe->getactivesheet()->setcellvalue('A' . $k, $field['COLUMN_NAME']);
        $obpe->getactivesheet()->setcellvalue('B' . $k, $field['COLUMN_TYPE']);
        $obpe->getactivesheet()->setcellvalue('C' . $k, $field['IS_NULLABLE']);
        $obpe->getActiveSheet()->mergeCells('D' . $k . ':F' . $k);
        $obpe->getactivesheet()->setcellvalue('D' . $k, $field['COLUMN_COMMENT']);
        $k++;
    }

    // 设置边框
    $obpe->getActiveSheet()->getStyle( 'A' . $first_key . ':F' . ($k - 1))->applyFromArray($styleThinBlackBorderOutline);

}


$obpe->setActiveSheetIndex(0);

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="mysql.xls"');
header('Cache-Control: max-age=0');

$objWriter = \PHPExcel_IOFactory::createWriter($obpe, 'Excel5');
$objWriter->save('php://output');
exit;
