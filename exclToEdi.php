<!DOCTYPE html>
<html>
<head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>Export Booking Excel to Coprar Converter</title>
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"
        integrity="sha256-4+XzXVhsDmqanXGHaHvgh1gMQKX40OUvDEBTu8JcmNs="
        crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/jszip.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.js"></script>
  </head>
<body>


<!--form method="post" enctype="multipart/form-data">
    <div class="form-group">
        <label for="exampleInputFile">File Upload</label>
        <input type="file" name="file" class="form-control" id="exampleInputFile">
    </div>
    <button type="submit" name="submit" class="btn btn-primary">Submit</button>
</form-->





<?php
require_once 'vendor/autoload.php';
// require_once 'db.php';
  
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
  
if (!isset($_POST['submit'])) 
{
    
    $html = '<div class="container">
        <form method="post" enctype="multipart/form-data" action="exclToEdi.php" onsubmit="return validateForm()" >
            <div class="card" style="">
                <div class="card-body">
                    <h5 class="card-title">Export Booking Excel to Coprar Converter</h5>
                    <div class="form-group">
                        <label for="recv_code">Receiver Code:</label><input class="form-control" type="text" id="recv_code" name="recv_code" value="RECEIVER" />
                        <p><small>Please change before file select.</small></p>
                    </div>
                    <div class="form-group">
                        <label for="recv_code">Callsign Code:</label><input class="form-control" type="text" id="callsign_code" name="callsign_code" value="XXXXX" />
                        <p><small>Please change before file select.</small></p>
                    </div>
                    <div class="form-group">
                        <label for="exampleInputFile">File Upload</label>
                        <input type="file" name="file" class="form-control" id="exampleInputFile">
                    </div>
                    
                    
                    <button type="submit" name="submit" class="btn btn-primary" style="float: right;">Submit</button>
                    <br>
                    <br>
                    <div class="form-group"><textarea class="form-control" rows="20" cols="40" id="my_file_output"></textarea></div>

                    
                </div>
            </div>
        </form>
    </div>';
    echo $html;
}
else
{
        
    $receiver_code = $_POST['recv_code'];
    $callsign_code = $_POST['callsign_code'];

    if(is_null($_POST['recv_code']))
    {
        echo "File format error";
        exit();
    }
    
 
    $file_mimes = array('text/x-comma-separated-values', 'text/comma-separated-values', 'application/octet-stream', 'application/vnd.ms-excel', 'application/x-csv', 'text/x-csv', 'text/csv', 'application/csv', 'application/excel', 'application/vnd.msexcel', 'text/plain', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
     
    if(isset($_FILES['file']['name']) && in_array($_FILES['file']['type'], $file_mimes)) 
    {
     
        $arr_file = explode('.', $_FILES['file']['name']);
        $extension = end($arr_file);
     
        if('csv' == $extension) {
            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
        } else {
            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        }
 
        $spreadsheet = $reader->load($_FILES['file']['tmp_name']);        
        $sheetCount = $spreadsheet->getSheetCount();

        for ($j = 0; $j < $sheetCount; $j++) 
        {
            $sheet = $spreadsheet->getSheet($j);
            $sheetData = $sheet->toArray();

            //$sheetData = $spreadsheet->getActiveSheet()->toArray();
            //todo - iterate through all worksheets
            $line = 0;
            $contcount = 0;
    
            if (!empty($sheetData)) 
            {
                
                                    
                $date = new DateTime("now", new DateTimeZone('Asia/Kuala_Lumpur'));
                $refno = $date->format('YmdHis');                    
                $edi = "UNB+UNOA:2+KMT+" . $receiver_code . "+" . $date->format('Ymd:Hi') . "+" . $refno . "'\n";
                $edi .= "UNH+" . $refno . "+COPRAR:D:00B:UN:SMDG21+LOADINGCOPRAR'\n"; 
                $line++;

                //1. - process header -
                $report_dt = ""; $voyage = ""; $vslname = ""; $callsign = ""; $opr = "";
            
                //1.1 get report date
                $tmpdt = explode("/",$sheetData[1][1]);
                $day = $tmpdt[0];
                $month = $tmpdt[1];
                $tmpyear = explode(" ",$tmpdt[2])[0];
                $tmptime = explode(" ",$tmpdt[2])[1];
                $report_dt = date('YmdHis', strtotime($tmpyear . "-" . $month . "-" . $day . " " . $tmptime));

                //1.2 get main report info
                if($sheetData[3][3] != "")
                {
                    $code = explode("/",$sheetData[3][3]);            
                    $voyage = $code[0]; 
                    $callsign = $code[1];
                    $opr = $code[2];
                    $vslname = $sheetData[3][1];
                }

                //1.3 include new info in edi
                $edi .= "BGM+45+" . $report_dt . "+5'\n"; 
                $line++;
                $edi .= "TDT+20+" . $voyage . "+1++172:" . $opr . "+++" . $callsign_code . ":103::" . $vslname . "'\n"; 
                $line++;
                $edi .= "RFF+VON:" . $voyage . "'\n"; 
                $line++;
                $edi .= "NAD+CA+" . $opr . "'\n"; 
                $line++; 

                for ($row=8; $row<count($sheetData); $row++) 
                {
                    $contcount++;

                    //rowCells[3] //5 - F, 4 - E
                    $fe = "5";              
                    if($sheetData[$row][3]=="E") 
                        $fe = "4";
                    else
                        $fe = "5"; 

                    //2 TS - N, 6 TS - Y                
                    if($sheetData[$row][11]=="Y") 
                        $type = "6";
                    else
                        $type = "2";

                    if(isset($sheetData[$row][1]) && isset($sheetData[$row][7]))
                        $edi .= "EQD+CN+" . $sheetData[$row][1] . "+" . $sheetData[$row][7] . ":102:5++" . $type . "+" . $fe . "'\n"; 
                    $line++; 

                    if(isset($sheetData[$row][5]))
                        $edi .= "LOC+11+" . $sheetData[$row][5] . ":139:6'\n"; 
                    $line++;

                    if(isset($sheetData[$row][6]))
                        $edi .= "LOC+7+" . $sheetData[$row][6] . ":139:6'\n";
                    $line++;

                    if(isset($sheetData[$row][19]))
                        $edi .= "LOC+9+" . $sheetData[$row][19] . ":139:6'\n";
                    $line++;
                    
                    if(isset($sheetData[$row][13]))
                        $edi .= "MEA+AAE+VGM+KGM:" . $sheetData[$row][13] . "'\n";
                    $line++;

                                    
                    if(isset($sheetData[$row][17]) && trim($sheetData[$row][17]," ") != "" && trim($sheetData[$row][17]," ") != "/") 
                    {
                        $tmp = explode(",",$sheetData[$row][17]);
                        for($i = 0; $i < count($tmp); $i++) 
                        {
                            $dim = explode("/",$sheetData[$row][17]);
                            if(trim($dim[0], " ")=="OF") 
                            {
                                $edi .= "DIM+5+CMT:" . trim($dim[1], " ") . "'\n"; 
                                $line++;
                            }
                            if(trim($dim[0]," ") == "OB") 
                            {
                                $edi .= "DIM+6+CMT:" . trim($dim[1]," ") . "'\n"; 
                                $line++;
                            }
                            if(trim($dim[0], " ") == "OR") 
                            {
                                $edi .= "DIM+7+CMT::" . trim($dim[1]," ") . "'\n"; 
                                $line++;
                            }
                            if(trim($dim[0], " ") == "OL") 
                            {
                                $edi .= "DIM+8+CMT::" . trim($dim[1]," ") . "'\n"; 
                                $line++;
                            }
                            if(trim($dim[0], " ") == "OH") 
                            {
                                $edi .= "DIM+9+CMT:::" . trim($dim[1]," ") . "'\n"; 
                                $line++;
                            }
                        }
                    
                    }

                
                    if(isset($sheetData[$row][15]) && trim($sheetData[$row][15]," ") != "" && trim($sheetData[$row][15]," ") != "/") 
                    {
                        $temperature = $sheetData[$row][15];
                        $temperature = str_replace(" ","", $temperature);
                        $temperature = str_replace("C","", $temperature);
                        $temperature = str_replace("+","", $temperature);                    
                        $edi .= "TMP+2+" . $temperature . ":CEL'\n"; 
                        $line++;
                    }

                    if(isset($sheetData[$row][25]) && trim($sheetData[$row][25]," ") != "" && trim($sheetData[$row][25]," ") != "/") 
                    {
                        $tmp = explode(",",trim($sheetData[$row][25]," "));
                        if($tmp[0]=="L") 
                        {
                            $edi .= "SEL+" . $tmp[1] . "+CA'\n"; 
                            $line++; //seal L - CA, S - SH, M - CU
                        }
                        if($tmp[0]=="S") 
                        {
                            $edi .= "SEL+" . $tmp[1] . "+SH'\n"; 
                            $line++; //seal L - CA, S - SH, M - CU
                        }
                        if($tmp[0]=="M") 
                        {
                            $edi .= "SEL+" . $tmp[1] . "+CU'\n"; 
                            $line++; //seal L - CA, S - SH, M - CU
                        }                  
                    }

                    if(isset($sheetData[$row][8])) 
                    {  
                        $edi .= "FTX+AAI+++" . $sheetData[$row][8]  . "'\n"; 
                        $line++; 
                    }                      
                    if(isset($sheetData[$row][12]) && trim($sheetData[$row][12]," ") != "" && trim($sheetData[$row][12]," ") != "/") 
                    {
                        $edi .= "FTX+AAA+++" . preg_replace('/[\x00-\x1F\x7F]/u', '', trim($sheetData[$row][12], " ")) . "'\n"; 
                        $line++;
                    }
                    if(isset($sheetData[$row][18]) && trim($sheetData[$row][18]," ") != "" && trim($sheetData[$row][18]," ") != "/")  
                    {
                        $edi .= "FTX+HAN++" . trim($sheetData[$row][18]," ") . "'\n";
                        $line++;
                    }
                    if(isset($sheetData[$row][14]) && trim($sheetData[$row][14]," ") != "" && trim($sheetData[$row][14]," ") != "/") 
                    {
                        $tmp = explode("/", $sheetData[$row][14]);
                        $edi .= "DGS+IMD+" . $tmp[0] . "+" . $tmp[1] . "'\n"; 
                        $line++;
                    }
                    if(isset($sheetData[$row][2]) && trim($sheetData[$row][2]," ") != "") 
                    { 
                        $edi .= "NAD+CF+" . trim($sheetData[$row][2]," ") . ":160:ZZZ'\n"; 
                        $line++; 
                    } //box
                        
                }
            }

        
            $contcount--;
            $edi .= "CNT+16:" . strval($contcount) . "'\n"; 
            $line++; 
            $line++;
            $edi .= "UNT+" . $line . "+" . $refno . "'\n";
            $edi .= "UNZ+1+" . $refno . "'";

        }
    }

    $html = '<div class="container">
                <form method="post" enctype="multipart/form-data" action="exclToEdi.php" onsubmit="return validateForm()">
                    <div class="card" style="">
                        <div class="card-body">
                            <h5 class="card-title">Export Booking Excel to Coprar Converter</h5>
                            <div class="form-group">
                                <label for="recv_code">Receiver Code:</label><input class="form-control" type="text" id="recv_code" name="recv_code" value="' . $receiver_code . '" />
                                <p><small>Please change before file select.</small></p>
                            </div>
                            <div class="form-group">
                                <label for="recv_code">Callsign Code:</label><input class="form-control" type="text" id="callsign_code" name="callsign_code" value="' . $callsign_code . '" />
                                <p><small>Please change before file select.</small></p>
                            </div>
                            <div class="form-group">
                                <label for="exampleInputFile">File Upload</label>
                                <input type="file" name="file" class="form-control" id="exampleInputFile">
                            </div>
                         
                            <button type="submit" name="submit" class="btn btn-primary" style="float: right;">Submit</button><br>
                            <br>
                            <div class="form-group"><textarea class="form-control" rows="20" cols="40" id="my_file_output">' . $edi . '</textarea></div>
                            
                        </div>
                    </div>
                </form>
            </div>';
            echo $html;
}

?>

<script>
    function validateForm()
    {
        if($('#recv_code').val()=="")
        {
            alert("Receiver Code cannot be empty");
            return false;
        }
        if($('#callsign_code').val()=="")
        {
            alert("Callsign Code cannot be empty");
            return false;
        }
        if($('#exampleInputFile').val()=="")
        {
            alert("Please select a file");
            return false;
        }
        return true;
    }
</script>

</body>
</html>




