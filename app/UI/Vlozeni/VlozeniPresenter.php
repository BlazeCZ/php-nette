<?php

declare(strict_types=1);

namespace App\UI\Vlozeni;

use Nette;
use PhpOffice\PhpSpreadsheet\IOFactory;


final class VlozeniPresenter extends Nette\Application\UI\Presenter
{

	public function __construct(
		private Nette\Database\Explorer $database,
	) {
	}
    

    public function actionUpload()
    {
        $file = $this->getHttpRequest()->getFile('file');
        
        if (!$file instanceof \Nette\Http\FileUpload) {
            $this->flashMessage('Chyba při nahrávání souboru.', 'error');
            $this->redirect('default');
        }
        
        if (!$file->isOk()) {
            $this->flashMessage('Chyba při nahrávání souboru.', 'error');
            $this->redirect('default');
        }

        if ($file->getContentType() !== 'application/vnd.ms-excel' && $file->getContentType() !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        $this->flashMessage('Soubor musí být ve formátu .xls nebo .xlsx.', 'error');
        $this->redirect('default');
        }

        try {
            $spreadsheet = IOFactory::load($file->getTemporaryFile());
            $worksheet = $spreadsheet->getActiveSheet();
            $rows = $worksheet->toArray();

            foreach ($rows as $index => $row) {
                if ($index === 0) {
                    continue; 
                }
                
                if (empty($row[0])) {
                    break;
                }
                
                if (!empty($row[3])) {
                 $amount = preg_replace('/\s+|Kč/', '', $row[3]);
                 $amount = str_replace(',', '', $amount);
                 $amountFormatted = round((float) $amount, 2);
                }
                
                if (!empty($row[2])) {
                 $closeDate = \DateTime::createFromFormat('m/d/Y', $row[2]);
                }
                
                $this->database->table('import_data')->insert([
                    'name' => $row[0],
                    'amount' => $amountFormatted,
                    'close_date' => $closeDate,
                    'status' => $row[4],
                    'description' => $row[5],
                    'offer' => $row[1],
                ]);
            
            }
            $this->flashMessage('Soubor byl úspěšně nahrán a data byla uložena do databáze.', 'success');
            
        } catch (\Exception $e) {
            $this->flashMessage('Chyba při zpracování souboru: ' . $e->getMessage(), 'error');
        }
        
        $this->redirect('default');
        
        
    }

            
            
}