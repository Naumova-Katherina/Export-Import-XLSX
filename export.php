public function actionPharmacyexport()
    {
        if(Yii::$app->request->isPost)
        {
            // данные приходят из POST запроса
            $pos = Yii::$app->request->post();
            $net = (int) $pos['net'];
            $holding = (int) $pos['holding'];
            $active = (int) $pos['active'];
            $str = $pos['str'];
            $prm = $pos['prm'];
        }

        // получаем данные из активных чекбоксов
        $check = array();
        foreach ($prm as $key => $value) { 
            foreach ($value as $key2 => $value2) {
                // берем все значения из массива
                $array = array_push($check, $value2);
                   if ($key2 == 'value' && $value2 != 'on' ) {
                       // удаляем значения не value => on из массива
                        array_pop($check);
                        array_pop($check);
                    } 
            }
        }
        
        //  удаление значений 'on' из массива $check
        $checkboxes = [];
        foreach ($check as $value) {
            if ($value != 'on') {
                array_push($checkboxes, $value);
            }
            // чекбокса claster_name нет
            if ($value == 'claster_id') {
                array_push($checkboxes, 'claster_name');
            }
            if ($value == 'subclaster_id') {
                array_push($checkboxes, 'subclaster_name');
            }
        }

        // построение запроса к БД
        $model;
        $s = ctype_digit($str) ? $str : '%' . $str . '%';

        // Переменные для подстановки в запрос
        
        $region = 'r.name as `region`';
        $net = 'n.name as `net`';
        $web = 'w.web';
        $subclaster_id = 's.id as `subclaster_id`';
        $subclaster = 's.name as `subclaster_name`';
        $claster_id = 's.claster_id as `claster_id`';
        $claster = 'cl.name as `claster_name`';
        $sub_report = 'sr.name as `subclaster_report`';
        $action = 'a.msg as `action`';
        $res_url = 'ru.url as `reservation_url`';
        
        // массив параметров из таблицы pharmacy
        $select = [];
        // массив параметров из JOIN таблиц
        $join = [];
        foreach ($checkboxes as $value) {
            if($value == "net_id"){
                $select[] = $net;
                $join[] = 'JOIN net n ON p.net_id = n.id ';
            } 
            elseif($value == "web_id"){
                $select[] = $web;
                $join[] = 'JOIN web_url w on p.web_id = w.id ';
            }
            elseif($value == "region_id"){
                $select[] = $region;
                $join[] = 'JOIN region_site r on p.region_id = r.id ';
            }
            elseif($value == "subclaster_id"){
                $select[] = $subclaster_id;
                $select[] = $subclaster;
                $join[] = 'JOIN subclaster s on p.subclaster_id = s.id ';
            }
            elseif($value == "claster_id"){
                $select[] = $claster_id;
                $select[] = $claster;
                $join[] = 'JOIN claster cl on s.claster_id = cl.id ';
            }
            
            elseif($value == "subclaster_report_id"){
                $select[] = $sub_report;
                $join[] = 'JOIN subclaster_report sr on p.subclaster_report_id = sr.id ';
            } 
            elseif($value == "action_id"){
                $select[] = $action;
                $join[] = 'JOIN action a on p.action_id = a.id ';
            }
            elseif($value == "reservation_url_id"){
                $select[] = $res_url;
                $join[] = 'JOIN resarvation_url ru on p.reservation_url_id = ru.id ';
            }
            elseif($value == "work_time"){
                $select[] = "p.is24";
                $select[] = "p.mon";
                $select[] = "p.tues";
                $select[] = "p.wedn";
                $select[] = "p.thur";
                $select[] = "p.fri";
                $select[] = "p.sat";
                $select[] = "p.san";
                $select[] = "p.remark";
            }
            elseif($value == "geo"){
                $select[] = "geo_x";
                $select[] = "geo_y";
                $select[] = "yandex_map";
                $select[] = "t_geo_x";
                $select[] = "t_geo_y";
            }
            elseif($value == "subclaster_name" or $value == "claster_name"){
                //не должны попасть в выборку!! таблица не pharmacy 
            }
            elseif ($value == "update_at"){
                $select[] = 'FROM_UNIXTIME(' . $value . ", '%Y-%m-%d %h:%i')";
            }
            else {                   
                $select[] = 'p.' . $value;
            }
        }
                  
        $where = '';
        if($net == 0 && $holding == 0)
        {
            $where .= ctype_digit($s) ? (' WHERE p.id = '.$s) :( ' WHERE p.name LIKE \''.$s.'\' OR p.address LIKE \''.$s.'\' OR p.identifier LIKE \''.$s.'\'');
        }
        else
        {
            $where .= ' WHERE ';
            $where .= $net == 0 ? '' : ' p.net_id = '.$net;
            $where .= $query == ' WHERE ' ? '' : ' AND ';
            $where .= $holding == 0 ? '' : 'n.holding_id = '.$holding;
        }
        if($active == 1){
            $where .= " AND p.available = 1";
        }

        // Тело запроса
        $query = "SELECT p.id, " . implode(', ',$select) . " FROM pharmacy p " . implode(' ', $join) . $where;  

        $model = Pharmacy::getDb()->createCommand($query)->queryAll();
        
        try {
            
            
        } catch (\Exception $e) 
        {
            echo $e->getMessage();
        }


        $date = date('Y-m-d H.i.s');
        $file = Yii::getAlias('@backend') . "/web/export/"."Аптеки - выгрузка.xlsx";
        
        $writer = new XLSXWriter();
        
        $name = 'Аптеки за '. $date;
            
        $headerStyle = array(
            'font-size'=>11,
            'font-style'=>'bold', 
            'halign'=>'center', 
            'border'=>'left,right,top,bottom',
            'border-style'=>'thin',
            'wrap_text'=>'true',
            'valign' => 'top',
            'fill' => '#f6e57b');
        $otherStyle = array(
            'font-size'=>10,
            'font'=>'Verdana',
            'halign'=>'left', 
            'border'=>'left,right,top,bottom',
            'border-style'=>'thin',
            'valign' => 'top');

        $sheet_name = $name . $date;
        
        $headerNames = [];
        // id не приходит с чекбоксами, необходимо добавить!
        $headerNames['id'] = 'id';

        $header = Pharmacy::getListOfFields();

        // получаем массив с заголовками столбцов
        foreach($header as $v) {
            foreach($v as $a) {
                foreach($a as $key=>$k) {
                    $headerNames[$key] = $k;
                }
            }
        }

        // массив для формата ячеек файла
        $format = [
            'id' => 'integer',
            'name' => 'string',
            'address' => 'string',
            'region_id'=> 'string',
            'net_id' => 'string',
            'web_id' => 'string',
            'phone' => 'string',
            'phone_int' => 'string',
            'subclaster_id' => 'integer',
            'subclaster_name' => 'string',
            'available' => 'integer',
            'create_at' => 'string',
            'update_at' => 'string',
            'notice' => 'string',
            'administrator' => 'string',
            'small_name' => 'string',
            'industrial' => 'integer',
            'state' => 'integer',
            'reservation_url_id' => 'string',
            'connected' => 'integer',
            'modify_at' => 'integer',
            'archived' => 'integer',
            'type_show_price' => 'integer',
            'type_show_amount' => 'integer',
            'is_use_net_type' => 'integer',
            'uid' => 'integer',
            'claster_id' => 'integer',
            'claster_name' => 'string',
            'action_id' => 'string',
            'subclaster_report_id' => 'string',
            'is24' => 'integer',
            'mon' => 'string',
            'tues' => 'string',
            'wedn' => 'string',
            'thur' => 'string',
            'fri' => 'string',
            'sat' => 'string',
            'san' => 'string',
            'remark' => 'string',
            'geo_x' => '#,##0.000000',
            'geo_y' => '#,##0.000000',
            'yandex_map' => 'string',
            't_geo_x' => '#,##0.000000',
            't_geo_y' => '#,##0.000000'
        ];

        //Добавим поля из группы Work Time к $checkboxes

        foreach ($checkboxes as $item) {
            if ($item == 'work_time') {
                $checkboxes[] = 'is24';
                $checkboxes[] = 'mon';
                $checkboxes[] = 'tues';
                $checkboxes[] = 'wedn';
                $checkboxes[] = 'thur';
                $checkboxes[] = 'fri';
                $checkboxes[] = 'sat';
                $checkboxes[] = 'san';
                $checkboxes[] = 'remark';
            }
        }

        //Добавим поля из группы Geo к $checkboxes

        foreach ($checkboxes as $item) {
            if ($item == 'geo') {
                $checkboxes[] = 'geo_x';
                $checkboxes[] = 'geo_y';
                $checkboxes[] = 'yandex_map';
                $checkboxes[] = 't_geo_x';
                $checkboxes[] = 't_geo_y';
            }
        }

        // получаем названия столбцов из массивов $checkboxes и $headerNames 
        foreach ($headerNames as $key => $value) {
            if ($key != 'id') {
                foreach ($checkboxes as $item) {
                    if ($item == $key) {
                        $names[$item] = $value;
                    }
                }
            }
        } 
        
        // в чекбоксах нет id!
        $add = array('id'=>'id'); 
        $names = $add + $names;

        // получаем массив названий чекбоксов $names => формат ячеек $format
        foreach ($format as $key => $value) {
            foreach ($names as $key2 => $value2) {
                if ($key == $key2) {
                    $result[$value2] = $value;
                }
            }      
        } 
        
        //  записываем первую строку с названиями столбцов, отдельные стили
        $writer->writeSheetHeader($sheet_name, $result, $headerStyle);

        //  записываем все остальные строки, данные в $model
        foreach($model as $record)
        {
            $writer->writeSheetRow($sheet_name, $record, $otherStyle);
        }

        // сохраняем XSLX файл по адресу на сервере
        $writer->writeToFile(Yii::getAlias('@backend') . "/web/export/"."Аптеки - выгрузка.xlsx");

       // отправляем в AJAX метод адрес для скачивания файла
        $url = "/export/Аптеки - выгрузка.xlsx";

        return $url; 
         
    }

    public function actionExportReservate() {
        //данные из формы
        if(Yii::$app->request->isPost)
        { 
            $post = Yii::$app->request->post(); 
        }

        $prm = $post['prm'];

        $date1 = '0';
        $date2 = '0'; 
        $apt = ''; 
        $ls = '';
        $email = '';
        $dateoff = '';

        // получаем данные из POST prm
        foreach ($prm as $item) { 
            foreach ($item as $key => $value) {
                // берем все значения из массива
                if ($key == 'name' && $value == 'apt') {
                    $array_apt = $item;
                }
                if ($key == 'name' && $value == 'ls' ) {
                    $array_ls = $item;
                }
                if ($key == 'name' && $value == 'email' ) {
                    $array_email = $item;
                }
                if ($key == 'name' && $value == 'date1' ) {
                    $array_date1 = $item;
                }
                if ($key == 'name' && $value == 'date2' ) {
                    $array_date2 = $item;
                }
                if ($key == 'name' && $value == 'dateoff' ) {
                    $dateoff = 1;
                }        
            }
        }

        // данные из формы получение значений
        foreach ($array_apt as $key => $value) {
            if ($key == 'value' && (is_string($value) || is_int($value))) {
                $apt = $value;
            }
        }

        foreach ($array_ls as $key => $value) {
            if ($key == 'value' && (is_string($value) || is_int($value))) {
                $ls = $value;
            }
        }

        foreach ($array_email as $key => $value) {
            if ($key == 'value' && (is_string($value) || is_int($value))) {
                $email = $value;
            }
        }

        if(isset($array_date1)) {
            foreach ($array_date1 as $key => $value) {
                if ($key == 'value') {
                    $date1 = $value;
                }
            }
        }

        if(isset($array_date2)) {
            foreach ($array_date2 as $key => $value) {
                if ($key == 'value') {
                    $date2 = $value;
                }
            }
        }

        // массив строк
        $model = Reservation::getReservation($apt, $ls, $email, $date1, $date2, $dateoff);

        // убираем лишнюю колонку id_hash, колонку status, редактирование status_exec
        for($i=0;$i<count($model);$i++) {
           unset($model[$i]["id_hash"]);
           if($model[$i]["status_exec"] == 1){
            $model[$i]["status_exec"] = 'Выполнено';
           }
           if($model[$i]["status_exec"] == 2){
            $model[$i]["status_exec"] = 'Отклонено';
           }
           if($model[$i]["status_exec"] == 0 && $model[$i]["status"] == 0){
            $model[$i]["status_exec"] = 'Не обработано';
           }
           if($model[$i]["status_exec"] == 0 && $model[$i]["status"] == 1){
            $model[$i]["status_exec"] = 'Аптекой получено';
           }
           unset($model[$i]["status"]);
        }
       

       // Yii::error($model);

        $date = date('Y-m-d H.i.s');
        $file = Yii::getAlias('@backend') . "/web/export/"."Бронь - выгрузка.xlsx";
        
        $writer = new XLSXWriter();
        
        $name = 'Бронь пользователей';
            
        //стили строк в файле
        $headerStyle = array(
            'font-size'=>11,
            'font-style'=>'bold', 
            'halign'=>'center', 
            'border'=>'left,right,top,bottom',
            'border-style'=>'thin',
            'wrap_text'=>'true',
            'valign' => 'top',
            'fill' => '#f6e57b');
        $otherStyle = array(
            'font-size'=>10,
            'font'=>'Verdana',
            'halign'=>'left', 
            'border'=>'left,right,top,bottom',
            'border-style'=>'thin',
            'valign' => 'top');

        $sheet_name = $name . $date;
            
        // массив заголовка в файле
        $header = [
            'order_id' => 'integer',
            'apt_id' => 'integer', 
            'name' => 'string',
            'address' => 'string',
            'ls_num' => 'integer',
            'ls_name' => 'string',
            'form' => 'string',
            'mnf' => 'string',
            'country' => 'string',
            'ls_price' => '#,##0.00',
            'amount' => 'integer',
            'barcode' => 'integer',
            'ls_apt' => 'string',
            'user_name' => 'string',
            'user_email' => 'string',
            'status_exec' => 'string',
            'date_add' => 'string',
            'date_exec' => 'string',
            'date_give' => 'string' 
        ];

        //  записываем первую строку с названиями столбцов, отдельные стили
        $writer->writeSheetHeader($sheet_name, $header, $headerStyle);

        //  записываем все остальные строки, данные в $model
        foreach($model as $record)
        {
            $writer->writeSheetRow($sheet_name, $record, $otherStyle);
        }

        // сохраняем XSLX файл по адресу на сервере
        $writer->writeToFile(Yii::getAlias('@backend') . "/web/export/"."Бронь - выгрузка.xlsx");

       // отправляем в AJAX метод адрес для скачивания файла
        $url = "/export/Бронь - выгрузка.xlsx";

        return $url;

    }
