# TP5_export_excel
TP5导出excel


```php
/**
 * 导出excel
 * @param  string $xlsTitle 导出文件名字
 * @param   array $expCellName 表头
 * @param   array $expTableData 导出数据
 * @return array        excel文件内容数组
 */
function export_excel($expTitle,$expCellName,$expTableData){
    $xlsTitle = iconv('utf-8', 'gb2312', $expTitle);//文件名称
    $fileName = $xlsTitle.'-'.date('YmdHis');//or $xlsTitle 文件名称可根据自己情况设定
    $cellNum = count($expCellName);
    $dataNum = count($expTableData);
    require ('../extend/phpexcel/Classes/PHPExcel.php');
    $objPHPExcel = new PHPExcel();
    $cellName = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ');

    for($i=0;$i<$cellNum;$i++){
        //填充颜色
        // $objPHPExcel->getActiveSheet(0)->getStyle($cellName[$i].'1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        // $objPHPExcel->getActiveSheet(0)->getStyle($cellName[$i].'1')->getFill()->getStartColor()->setARGB('FFccc');
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($cellName[$i].'1', $expCellName[$i][1]);
        $objPHPExcel->getActiveSheet()->getStyle($cellName[$i].'1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $objPHPExcel->getActiveSheet()->getStyle($cellName[$i].'1')->getFill()->getStartColor()->setARGB('FFFFB400');
        $k=$i+1;
        $objPHPExcel->getActiveSheet()->getRowDimension($k)->setRowHeight(25);
        
    }

    // Miscellaneous glyphs, UTF-8
    for($i=0;$i<$dataNum;$i++){
        for($j=0;$j<$cellNum;$j++){
            $objPHPExcel->getActiveSheet(0)->setCellValue($cellName[$j].($i+2), $expTableData[$i][$expCellName[$j][0]]);
            $objPHPExcel->getActiveSheet(0)->getColumnDimension($cellName[$j])->setAutoSize(true);
        }
    }
    // 输出Excel表格到浏览器下载
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="'.$fileName.'.xls"');
    header('Cache-Control: max-age=0');
    // If you're serving to IE 9, then the following may be needed
    header('Cache-Control: max-age=1');
    // If you're serving to IE over SSL, then the following may be needed
    header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
    header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
    header('Pragma: public'); // HTTP/1.0


    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $temp = 'Uploads/temp/'.$fileName.".xls";
    //$objWriter->save($temp);
    $objWriter->save('php://output');
    //return $temp;
}



//导出商品订单
    public function delivery_export(){
        
        $ordername = input('param.ordername');
        if(input('param.order_type_id')!=null){
            $where['order_type_id'] = input('param.order_type_id');
        }

        if(input('param.yuding') !=null){
            $where['yuding'] = input('param.yuding');
        }

        if (input('param.status') !=null) {
           $where['status'] = input('param.status');
        }

        if(input('param.type') != ''){
            $where['status'] = input('param.type');
        }
        if(input('param.id') != ''){
            $where['order_id'] = input('param.id');
        }
        if(input('param.to') != ''){
            $where['create_time']=array(array('gt',input('param.to')),array('lt',input('param.so')));
        }

        $order_list = OrderGoods::where($where)
                        ->order('id desc')
                        ->select();

        $xq=array();
        foreach ($order_list as $k => $v) {
            $xq[] = DB::name('order_goods_middle')->where('order_id','eq',$v['order_id'])->select();
            foreach ($xq as $ks => $vs) {
                $a = array();
                foreach ($vs as $kss => $vss) {
                   $a[] = ' 商品编号：'.$vss['goods_id'].' 商品名称：'.$vss['goods_name'].' 单价：'.$vss['price'].' 数量：'.$vss['number'];
                }
            }
            $order_list[$k]['xiangq']=implode(',', $a);
            switch ($v['status']) {
                case '-1':
                    $order_list[$k]['status'] = '已取消';
                    break;
                case '0':
                    $order_list[$k]['status'] = '侍付款';
                    break;
                case '1':
                    $order_list[$k]['status'] = '已付款';
                    break;
                case '2':
                    $order_list[$k]['status'] = '待配送';
                    break;
                case '3':
                    $order_list[$k]['status'] = '配送中';
                    break;
                case '4':
                    $order_list[$k]['status'] = '已收货';
                    break;
                case '5':
                    $order_list[$k]['status'] = '已评价';
                    break;
                case '6':
                    $order_list[$k]['status'] = '预订中';
                    break;
                default:
                    # code...
                    break;
            }
        }
        
        $xlsCell  = array(
                array('id','编号'),
                array('order_id','订单编号'),
                array('realname','联系人'),
                array('phone','联系电话'),
                array('price','价格'),
                array('status','订单状态'),
                array('xiangq','订单详情'),
                array('create_time','下单时间')
                );
       // export_to($data,'用户列表');//导出excle
       export_excel($ordername,$xlsCell,$order_list);
   }
```









