<?php
  
  //Caminho do arquivo excel que será extraido os dados
  $filename = 'dados.xls';
   
   // If you need to parse XLS files, include php-excel-reader
	require('spreadsheet-reader/php-excel-reader/excel_reader2.php');

	require('spreadsheet-reader/SpreadsheetReader.php');

    $Reader = new SpreadsheetReader($filename);
    foreach ($Reader as $key => $value)
	{
        if($key > 1) {
        
        $agencia = $value[0];
        $descricao = $value[1];
        $nroContrato = $value[2];
        $periodo = $value[3];
        $documentoSD = $value[4];
        $dataRetirada = $value[5];
        $dataDevolucao = $value[6];
        $nomeFirma = $value[7];
        $baseCalculo = $value[8];
        $comissao = $value[9];
        $moeda = $value[10];
        $vencimento = $value[11];
        $docCompra = $value[14];
        $nomeUsuario = $value[19];
        $voucher = $value[24];
        
        $result = "insert into dados(Status,AgenciaID,DescricaoAgencia,
        Voucher,NomeCliente,NroContrato,Periodo,DocSD,DataRetirada,DataDevolucao,NomeEmpresa
        BaseCalculo,Comissao,Moeda,Vencimento,DocCompra)
        
        values('','$agencia','$descricao','$voucher','$nomeUsuario','$nroContrato',
        '$periodo','$documentoSD','$dataRetirada','$dataDevolucao','$nomeFirma',
        '$baseCalculo','$comissao','$moeda','$vencimento','$docCompra')";
      

        }
    }
     if($result){
        echo 'Exportado relatorio com sucesso';
        }else{
            echo 'Error ao exportar o relatorio';
        }
     
 ?>
