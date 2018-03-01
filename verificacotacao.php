<?php
/**
 * Created by PhpStorm.
 * User: Elton
 * Date: 27/02/2018
 * Time: 16:24
 */

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

try {
    // Permiti a entrada de dados mais fácil
    PhpOffice\PhpSpreadsheet\Cell\Cell::setValueBinder(new \PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder());

    // Abre o arquivo
    $spreadsheet = IOFactory::load('CAMBIO.xlsx');

    // Retorna a guia ativa
    $sheet = $spreadsheet->getSheet(0);

    // Insere uma linha antes da linha 4
    $sheet->insertNewRowBefore(4, 1);

    // Copiar conteúdo
    $copiarConteudo = $sheet->rangeToArray('A3:G3', null, true, false);

    // Insere o conteúdo copiado na linha nova
    $sheet->fromArray($copiarConteudo, null, 'A4');

    // A data de hoje
    $dateTimeNow = time();

    // Conversão para um tipo compatível para o Excel
    $excelDataValue = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($dateTimeNow);

    // Acessa a API de cotação
    $json = acessarDadosCotacaoMoedas();

    // Acessa a página para ver o valor do ouro mil
    $ouro = acessarDadosCotacaoOuroMil();

    // Insere os novos valores
    $sheet->setCellValue('A3', $excelDataValue);
    $sheet->setCellValue('B3', $excelDataValue);
    $sheet->setCellValue('C3', $ouro);
    $sheet->setCellValue('D3', $json->USDT->bid);
    $sheet->setCellValue('E3', $json->USDT->ask);
    $sheet->setCellValue('F3', $json->EUR->bid);
    $sheet->setCellValue('G3', $json->EUR->ask);

    // Instância a planilha para salvar
    $writer = new Xlsx($spreadsheet);

    // Salva a planilha editada
    $writer->save('CAMBIO.xlsx');

    echo "Pronto!!";
} catch (Exception $e) {
    echo $e->getMessage();
}

/**
 * Acessa a url https://economia.awesomeapi.com.br/json/all para pegar os dados de cotação. É verificado se
 * as moedas dólar turismo e euro estão no retorno através do respectivo codein. E os valores são
 * transformados em float
 *
 * @return stdClass
 * @throws Exception
 */
function acessarDadosCotacaoMoedas()
{
    $site = file_get_contents('https://economia.awesomeapi.com.br/json/all');
    $json = json_decode($site);

    if ($json->USDT->codein !== 'BRLT') {
        throw new Exception('Nao foi retornado os dados do dolar turismo');
    }

    if ($json->EUR->codein !== 'BRL') {
        throw new Exception('Nao foi retornado os dados do euro');
    }

    try {
        $json->USDT->bid = (float) $json->USDT->bid;
        $json->USDT->ask = (float) $json->USDT->ask;
        $json->EUR->bid = (float) $json->EUR->bid;
        $json->EUR->ask = (float) $json->EUR->ask;
    } catch (Exception $e) {
        throw new Exception('Nao foi possivel converter os valores para float.');
    }

    return $json;
}

/**
 * Acessa a url https://dolarhoje.com/ouro-hoje/ para pegar os dados do ouro mil. Com os dados recebidos
 * é feito uma captura dos dados pertinentes e convertido para float o valor do ouro mil por 1g
 *
 * @return float
 * @throws Exception
 */
function acessarDadosCotacaoOuroMil()
{
    $site = file_get_contents('https://dolarhoje.com/ouro-hoje/');
    preg_match_all('/R\$\s([\d|,]*)<\/td/', $site, $saida);

    try {
        return (float) str_replace(',', '.', $saida[1][0]);
    } catch (Exception $e) {
        throw new Exception('Nao foi possivel verificar a cotacao do ouro mil');
    }
}