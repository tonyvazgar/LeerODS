package abrirods;
import java.io.File;
import java.io.IOException;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Range;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

public class AbrirODS {
    public void leer(File file){
        Sheet sheet;
        try{
            sheet = SpreadSheet.createFromFile(file).getSheet(0);
            
            int nColCount = 11;
            int nRowCount = 9;
            
            System.out.println("Rows: " + nRowCount + " Cols: " + nColCount);
            MutableCell cell = null;
            for(int nRowIndex = 2; nRowIndex < nRowCount; nRowIndex++){
               for(int nColIndex = 0 ;nColIndex < nColCount; nColIndex++){
                    cell = sheet.getCellAt(nColIndex, nRowIndex);
                    //System.out.print(cell.getValue()+ " | ");
                    switch(nColIndex){
                            case 0:
                                //document.append("NombrePeriodico", cell.toString());
                                System.out.print("Nombre Periodico: " + cell.getValue() + " | ");
                                break;
                            case 1:
                                //document.append("Número", cell.toString());
                                System.out.print("#: " + cell.getValue()+ " | ");
                                break;
                            case 2:
                                //document.append("Fecha", cell.toString());
                                System.out.print("Fecha: " + cell.getValue()+ " | ");
                                break;
                            case 3:
                                //document.append("Evento", cell.toString());
                                System.out.print("Evento: " + cell.getValue()+ " | ");
                                break;
                            case 4:
                                //document.append("Latitud", cell.toString());
                                System.out.print("Latitud: " + cell.getValue()+ " | ");
                                break;
                            case 5:
                                //document.append("Longitud", cell.toString());
                                System.out.print("Longitud: " + cell.getValue()+ " | ");
                                break;
                            case 6:
                                //document.append("VelocidadViento", cell.toString());
                                System.out.print("VelViento: " + cell.getValue()+ " | ");
                                break;
                            case 7:
                                //document.append("Creciente", cell.toString());
                                System.out.print("Creciente: " + cell.getValue()+ " | ");
                                break;
                            case 8:
                                //document.append("Temperatura", cell.toString());
                                System.out.print("Temp: " + cell.getValue()+ " | ");
                                break;
                            case 9:
                                //document.append("EfectoDetectado", cell.toString());
                                System.out.print("EfectoDetec: " + cell.getValue()+ " | ");
                                break;
                            case 10:
                                //document.append("URL", cell.toString());
                                System.out.print("URL: " + cell.getValue()+ " | ");
                                break;
                        }
                }
                System.out.println();
            }
        }catch(IOException e){
            e.printStackTrace();
        }
    }
    public static void main(String[] args) {
        File file = new File("C:\\Users\\tonyv\\Desktop\\1882.ods");
        AbrirODS objODSReader = new AbrirODS();
        objODSReader.leer(file);
    }
}