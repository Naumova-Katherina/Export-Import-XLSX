 public function actionImport() {
        $filemodel = new UploadImport();

        if (Yii::$app->request->isPost) {

            $filemodel->file = UploadedFile::getInstance($filemodel, 'file');

            if ($filemodel->file && $filemodel->validate()) {

                // если файл не прошел валидацию в модели UploadImport, то он не сохранится!!
                $filemodel->file->saveAs(Yii::getAlias('@app/web/uploads/prices/') . $filemodel->file->baseName . '.' . $filemodel->file->extension);
            } else {
                Yii::error('валидация модели НЕ прошла!');
                echo ('Валидация не прошла! Добавьте файл! Если файл был, то добавьте свой тип MIME в список!');
                die();
            }

            // загружаем xlsx file 
            if ($xlsx = SimpleXLSX::parse(Yii::getAlias('@app/web/uploads/prices/') . $filemodel->file->baseName . '.' . $filemodel->file->extension)) {

                $success = ('Файл успешно загружен!');
            } else {
                echo (SimpleXLSX::parseError());
                die();
            }
            // получаем массив строк таблицы exсel     
            $array = $xlsx->rows();

            // получаем строку с названиями колонок
            $names = $array[0];

            // ищем нужные названия колонок в первой строке
            foreach ($names as $key => $value) {
                if ($value == 'id') {
                    (int) $id = $key;
                }
                if ($value == 'uid') {
                    (int) $uid = $key;
                }
                if ($value == 'Субкластер') {
                    (int) $subclaster_id = $key;
                }
                if ($value == 'Субкластер Название') {
                    (string) $subclaster = $key;
                }
                if ($value == 'Кластер') {
                    (int) $claster_id = $key;
                }
                if ($value == 'Кластер Название') {
                    (string) $claster = $key;
                }
            }

            //cчетчик затронутых строк
            $count = 0;

            $success = '';

            // обновление только по ID => UID
            if (isset($uid)) {

                $success = $success . "  ID => UID  ";

                // получаем массив из значений ID => Uid из файла
                $column_uid = ArrayHelper::map($array, $id, $uid);

                //метод обновления по id=>uid
                $count_uid = 0;
                $count_uid = Pharmacy::updateUid($column_uid);
            }

            // если в массиве существует claster_id, возможность его обновления
            if (isset($claster_id)) {

                $success = $success . "  ID => claster_id  ";

                $count_claster = 0;

                $count_claster = Claster_db_pharm::updateClaster($array, $id, $claster_id, $claster);
            }

            //обновление только по ID => subclaster_id
            if (isset($subclaster_id)) {

                $success = $success . "  ID => subclaster_id  ";
                $count_subclaster = 0;

                //метод обновления по id=>subclaster_id
                $count_subclaster = Subclaster_db_pharm::updateSubclaster($array, $id, $subclaster_id, $subclaster);
            }

            //удаляем загруженный файл
            if ($delete = unlink(Yii::getAlias('@app/web/uploads/prices/') . $filemodel->file->baseName . '.' . $filemodel->file->extension)) {
                $del = ('Файл успешно удален после обновления!');
            } else {
                $del = ('Файл не удалился, что-то пошло не так!');
            }
        }

        //общее число затронутых строк
        if (isset($count_uid)) {
            $count = $count + $count_uid;
        }

        if (isset($count_subclaster)) {
            $count = $count + $count_subclaster;
        }

        return $this->render('import', compact('filemodel', 'count', 'count_uid', 'count_subclaster', 'del', 'success', 'info'));
    }
