<?php

declare(strict_types=1);

namespace xhulav\SpreadsheetResponse;

use Nette\Application\Response;
use Nette\Http\IRequest;
use Nette\Http\IResponse;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class SpreadsheetResponse implements Response
{
	private string $fileName;
	private Spreadsheet $spreadsheet;

	public function __construct(Spreadsheet $spreadsheet, string $fileName = 'spreadsheet.xlsx')
	{
		$this->fileName = $fileName;
		$this->spreadsheet = $spreadsheet;
	}

	function send(IRequest $httpRequest, IResponse $httpResponse): void
	{
		// Set headers
		$httpResponse->setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		$httpResponse->setHeader('Content-Disposition', "attachment; filename={$this->fileName}");

		// Send the spreadsheet file
		IOFactory::createWriter($this->spreadsheet, 'Xlsx')
			->save('php://output');
	}
}