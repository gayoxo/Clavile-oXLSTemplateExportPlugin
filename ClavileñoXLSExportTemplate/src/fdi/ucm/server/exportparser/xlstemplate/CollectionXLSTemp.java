/**
 * 
 */
package fdi.ucm.server.exportparser.xlstemplate;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Random;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.CompleteLogAndUpdates;
import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteFile;
import fdi.ucm.server.modelComplete.collection.document.CompleteLinkElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElementFile;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElementURL;
import fdi.ucm.server.modelComplete.collection.document.CompleteTextElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteLinkElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteResourceElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteStructure;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteTextElementType;

/**
 * @author Joaquin Gayoso-Cabada
 *Clase qie produce el XLSI
 */
public class CollectionXLSTemp {


	public static String processCompleteCollection(CompleteLogAndUpdates cL,
			CompleteCollection salvar, boolean soloEstructura, String pathTemporalFiles) throws IOException {
		
		 /*La ruta donde se creará el archivo*/
        String rutaArchivo = pathTemporalFiles+"/"+System.nanoTime()+".xls";
        /*Se crea el objeto de tipo File con la ruta del archivo*/
        File archivoXLS = new File(rutaArchivo);
        /*Si el archivo existe se elimina*/
        if(archivoXLS.exists()) archivoXLS.delete();
        /*Se crea el archivo*/
        archivoXLS.createNewFile();
        
        /*Se crea el libro de excel usando el objeto de tipo Workbook*/
        Workbook libro = new HSSFWorkbook();
        
        /*Se inicializa el flujo de datos con el archivo xls*/
        FileOutputStream archivo = new FileOutputStream(archivoXLS);
        
        /*Utilizamos la clase Sheet para crear una nueva hoja de trabajo dentro del libro que creamos anteriormente*/
        
        HashMap<Long, Integer> clave=new HashMap<Long, Integer>();	
        
//        Sheet hoja;
        HashMap<CompleteGrammar, Sheet> TablaAmbitos=new HashMap<CompleteGrammar, Sheet>();
        HashMap<CompleteGrammar, Sheet> TablaDatos=new HashMap<CompleteGrammar, Sheet>();
        HashMap<CompleteGrammar, Integer> TablaNoname=new HashMap<CompleteGrammar, Integer>();

        int Value=0;
        
        
        
        for (CompleteGrammar row : salvar.getMetamodelGrammar()) {
        	   Sheet hoja;
        	   
        	   if (!row.getNombre().isEmpty())
        	   	 {
        				int indice = 0;
        				String Nombreactual=row.getNombre();
        				while (libro.getSheet(Nombreactual)!=null)
        					{
        					Nombreactual=row.getNombre()+indice;
        					indice++;
        					}
        				hoja = libro.createSheet(Nombreactual);
        	   	 }
        	   else{
        		   hoja = libro.createSheet();
        		   TablaNoname.put(row, Value);
      	        	Value++;
        	   }
       		TablaDatos.put(row, hoja);
		}

        
        for (CompleteGrammar row : salvar.getMetamodelGrammar()) {
        	Sheet hojaP = TablaDatos.get(row);
        	   Sheet hoja;
     	        	 hoja = libro.createSheet(hojaP.getSheetName()+"_Scopes");
        	   TablaAmbitos.put(row, hoja);
		}
        
        
        for (CompleteGrammar row : salvar.getMetamodelGrammar()) {
        	Sheet HojaDatos=TablaDatos.get(row);
        	Sheet HojaAmbitos=TablaAmbitos.get(row);
			processGrammar(libro,HojaDatos,HojaAmbitos,row,clave,cL,salvar.getEstructuras(),soloEstructura);
		}
        
        
//        /*Escribimos en el libro*/
        libro.write(archivo);
        /*Cerramos el flujo de datos*/
        archivo.close();
        /*Y abrimos el archivo con la clase Desktop*/
//        Desktop.getDesktop().open(archivoXLS);
		return rutaArchivo;
//        }
//        else 
//        	{
//        	 libro.write(archivo);
//        	archivo.close();
//        	return "";
//        	}
        

        
        
	}
	
	

	private static void processGrammar(Workbook libro, Sheet hojaDatos, Sheet hojaAmbitos, CompleteGrammar grammar,
			HashMap<Long, Integer> clave, CompleteLogAndUpdates cL, List<CompleteDocuments> list, boolean soloEstructura) {
		  
	  
	        
	        List<CompleteElementType> ListaElementos=generaLista(grammar);
	        

	        if (ListaElementos.size()>255)
	        	{
	        	cL.getLogLines().add("Tamaño de estructura demasiado grande para exportar a xls para gramatica: " + grammar.getNombre() +" solo 255 estructuras seran grabadas, divide en gramaticas mas simples");
	        	ListaElementos=ListaElementos.subList(0, 254);
	        	}
	        
	        List<CompleteDocuments> ListaDocumentos=generaDocs(list,grammar);
	      
	        if (ListaDocumentos.size()+1>65536)
	    	{
	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls");
	    	ListaDocumentos=ListaDocumentos.subList(0, 65536);
	    	}

	        	
	        int row=0;
	        int Column=0;
	        int columnsMax=ListaElementos.size();
	       	
	        
	        for (int i = 0; i < 2; i++) {
	        	Row filaDatos = hojaDatos.createRow(row);
	        	Row filaAmbitos = hojaAmbitos.createRow(row);
	        	row++;
	        	
	        	if (i==0)
	        	{
	        	for (int j = 0; j < columnsMax+2; j++) {
	        		
	        		String ValueDatos = "";
	            	if (j==0)
	            		ValueDatos="Column Clavy Document Id ( ADD NEGATIVE NUMBERS FOR NEW DOCS ) ";
	            	else 
	            		if (j==1)
	            			ValueDatos="Description";
	            		else
	            		{
	            		CompleteElementType TmpEle = ListaElementos.get(j-2);
	            		ValueDatos=pathFather(TmpEle);
	            		}
	
	            	
	            	if (ValueDatos.length()>=32767)
	            	{
	            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + ValueDatos + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		ValueDatos.substring(0, 32766);
	            	}
	            		Cell celdaDatos = filaDatos.createCell(j);
	            		Cell celdaAmbitos = filaAmbitos.createCell(j);
	            		
	            	if (j>1)
	            		{
	            		clave.put(ListaElementos.get(j-2).getClavilenoid(), Column);
	            		Column++;
	            		}
	            	else
	            	if (j==0)
	            	{
	            		hojaDatos.setColumnWidth(j, 15750);
	            	}
	            	else
	            	if (j==1)
		            {
	            		hojaDatos.setColumnWidth(j, 12750);
		            }
	            	
	            	celdaDatos.setCellValue(ValueDatos);
	            	
	            	if (j==0)
	            		celdaAmbitos.setCellValue("Scopes Values for import/export, be carefull if you modify this");
	            	
	           }
			}else if (i==1)
        	{
        		for (int j = 0; j < columnsMax+2; j++) {
	        		
	        		String ValueDatos = "";
	        		if (j==0)
	            		ValueDatos="Row Clavy Type Id ( DO NOT MODIFY THIS ROW )";
	            	else 
	            		if (j==1)
	            			ValueDatos=Long.toString(grammar.getClavilenoid());
	            		else
	            		{
	            		CompleteElementType TmpEle = ListaElementos.get(j-2);
	            		ValueDatos=Long.toString(TmpEle.getClavilenoid());
	            		}
	
	            	
	        		if (ValueDatos.length()>=32767)
	            	{
	            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + ValueDatos + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		ValueDatos.substring(0, 32766);
	            	}
	            		Cell celdaDatos = filaDatos.createCell(j);
	            		Cell celdaAmbitos=filaAmbitos.createCell(j);
	            	
	            	celdaDatos.setCellValue(ValueDatos);
	            	
	            	if (j==0)
	            		celdaAmbitos.setCellValue("Scopes Row/Colum match with the named colum Row/colum value");
	           }
        		
        		
        	}
	        	
	        }
	        
	        
	        if (!soloEstructura)
	        {
	        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
	        for(int f=0;f<ListaDocumentos.size();f++){
	            /*La clase Row nos permitirá crear las filas*/
	            Row filaDatos = hojaDatos.createRow(row);
	            Row filaAmbitos = hojaAmbitos.createRow(row);
	            row++;

	            CompleteDocuments Doc=ListaDocumentos.get(f);
	            HashMap<Integer, ArrayList<CompleteElement>> ListaClave=new HashMap<Integer, ArrayList<CompleteElement>>();
	            
	            for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					if (val!=null)
						{
						ArrayList<CompleteElement> Lis=ListaClave.get(val);
						if (Lis==null)
							{
							Lis=new ArrayList<CompleteElement>();
							}
						Lis.add(elem);
						ListaClave.put(val, Lis);
						}
				}
	            
	            
	            
	            /*Cada fila tendrá celdas de datos*/
	            for(int c=0;c<columnsMax+2;c++){
	            	
	            	String ValueDatos = "";
	            	String ValueAmbitos = "";
	            	if (c==0)
	            		ValueDatos=Long.toString(Doc.getClavilenoid());
	            	else if (c==1)
	            		ValueDatos=Doc.getDescriptionText();
	            	else
	            		{
	            		ArrayList<CompleteElement> temp = ListaClave.get(c-2);
	            		if (temp!=null)
	            		{
	            			
	            		//Calculo Value normal	
	            		if (temp.size()>1)
	            			{
	            			StringBuffer SB=new StringBuffer();
	            			SB.append("Size: " +temp.size());
	            			for (CompleteElement completeElement : temp) {
	            				SB.append("{");
	            				SB.append(getValueFromElement(completeElement));
								SB.append("}");
							}
	            			ValueDatos=SB.toString();
	            			}
	            		else if (temp.size()>0){
	            			CompleteElement completeElement=temp.get(0);
	            			if (completeElement instanceof CompleteTextElement)
		            			ValueDatos=(((CompleteTextElement)completeElement).getValue());
							else if (completeElement instanceof CompleteLinkElement)
							{
								try {
									ValueDatos=Long.toString((((CompleteLinkElement)completeElement).getValue().getClavilenoid()));
								} catch (Exception e) {
									ValueDatos="";
								}
		            			
							}
							else if (completeElement instanceof CompleteResourceElementURL)
		            			ValueDatos=(((CompleteResourceElementURL)completeElement).getValue());
							else if (completeElement instanceof CompleteResourceElementFile)
		            			ValueDatos=(((CompleteResourceElementFile)completeElement).getValue().getPath());
	            		}

	            		//Calculo value ambitos
	            		if (temp.size()>1)
            			{
            			StringBuffer SB=new StringBuffer();
            			SB.append("Size: " +temp.size());
            			for (CompleteElement completeElement : temp) {
            				SB.append("{");
            					SB.append(procesaAmbitos(completeElement.getAmbitos()));
            				
							SB.append("}");
						}
            			ValueAmbitos=SB.toString();
            			}
            		else if (temp.size()>0){
            			CompleteElement completeElement=temp.get(0);
            			ValueAmbitos=procesaAmbitos(completeElement.getAmbitos());
					
            		}
	            		
	            		
	            		
	            		}
	            		}
	
	            	 
	            	if (ValueDatos.length()>=32767)
	            	{
	            		ValueDatos="";
	            		cL.getLogLines().add("Tamaño de Texto en Valor en elemento " + ValueDatos + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		ValueDatos.substring(0, 32766);
	            	}
	                /*Creamos la celda a partir de la fila actual*/
	                Cell celdaDatos = filaDatos.createCell(c);
	                Cell celdaAmbitos = filaAmbitos.createCell(c);
	                
	                		 celdaDatos.setCellValue(ValueDatos);
	                		 celdaAmbitos.setCellValue(ValueAmbitos);
	                    /*Si no es la primera fila establecemos un valor*/

	                
	            	}

	            		
	            		
	            }
	        
	        }
	        
	       
		
	}

	private static String procesaAmbitos(ArrayList<Integer> ambitos) {
		StringBuffer SB=new StringBuffer();
		for (Integer integer : ambitos) {
				SB.append("{");
					SB.append(integer);
				
				SB.append("}");
			}
		return SB.toString();
	}



	private static ArrayList<CompleteDocuments> generaDocs(
			List<CompleteDocuments> list, CompleteGrammar grammar) {
		ArrayList<CompleteDocuments> ListaDoc=new ArrayList<CompleteDocuments>();
		for (CompleteDocuments completeDocuments : list) {
			if (StaticFuctionsXLSTemp.isInGrammar(completeDocuments,grammar))
				ListaDoc.add(completeDocuments);
		}
		return ListaDoc;
	}

//	private static ArrayList<CompleteElementType> generaLista(
//			List<CompleteGrammar> metamodelGrammar) {
//		  ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
//		  for (CompleteGrammar completegramar : metamodelGrammar) {
//			ListaElementos.addAll(generaLista(completegramar));
//		}
//		return ListaElementos;
//	}

	private static ArrayList<CompleteElementType> generaLista(
			CompleteGrammar completegramar) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 for (CompleteStructure completeelem : completegramar.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType)
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
				ListaElementos.addAll(generaLista(completeelem));
			}
		 return ListaElementos;
	}

	private static Collection<? extends CompleteElementType> generaLista(
			CompleteStructure completeelementPadre) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 for (CompleteStructure completeelem : completeelementPadre.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType)
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
				ListaElementos.addAll(generaLista(completeelem));
			}
		 return ListaElementos;
	}

	private static String getValueFromElement(CompleteElement completeElement) {
		try {
			if (completeElement instanceof CompleteTextElement)
    			return (((CompleteTextElement)completeElement).getValue());
			else if (completeElement instanceof CompleteLinkElement)
				return Long.toString((((CompleteLinkElement)completeElement).getValue().getClavilenoid()));
			else if (completeElement instanceof CompleteResourceElementURL)
				return (((CompleteResourceElementURL)completeElement).getValue());
			else if (completeElement instanceof CompleteResourceElementFile)
				return (((CompleteResourceElementFile)completeElement).getValue().getPath());
		} catch (Exception e) {
			return "";
		}
		return "";
	}
	
	public static void main(String[] args) throws Exception{
		
		int id=0;
		
		
		
		  CompleteCollection CC=new CompleteCollection("Lou Arreglate", "Arreglate ya!");
		  for (int i = 0; i < 5; i++) {
			  CompleteGrammar G1 = new CompleteGrammar(new Long(id),"Grammar"+i, i+"", CC);
			  
			  ArrayList<CompleteDocuments> CD=new ArrayList<CompleteDocuments>();
			  int docsN=(new Random()).nextInt(5);
			  docsN=docsN+5;
			for (int j = 0; j < docsN; j++) {
				CompleteDocuments CDDD=new CompleteDocuments(new Long(id), CC, "", "");
				CC.getEstructuras().add(CDDD);
				 id++;
				CD.add(CDDD);
			}
			  
			  id++;
			  for (int j = 0; j < 5; j++) {
				  CompleteElementType CX = new CompleteElementType(new Long(id),"Structure "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
			}
			  for (int j = 0; j < 5; j++) {
				  CompleteTextElementType CX = new CompleteTextElementType(new Long(id),"Texto "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
				
				for (CompleteDocuments completeDocuments : CD) {
					boolean docrep=(new Random()).nextBoolean();
					if (docrep)
						{
						CompleteTextElement CTE=new CompleteTextElement(new Long(id), CX, "Texto "+(i*10+j));
						id++;
						completeDocuments.getDescription().add(CTE);
						}
				}
				
				
				
			}
			  for (int j = 0; j < 5; j++) {
				  CompleteLinkElementType CX = new CompleteLinkElementType(new Long(id),"Link "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
				
				for (CompleteDocuments completeDocuments : CD) {
					boolean docrep=(new Random()).nextBoolean();
					if (docrep)
						{
						CompleteLinkElement CTE=new CompleteLinkElement(new Long(id), CX, CD.get((new Random()).nextInt(CD.size())));
						id++;
						completeDocuments.getDescription().add(CTE);
						}
				}
			}
			  for (int j = 0; j < 5; j++) {
				  CompleteResourceElementType CX = new CompleteResourceElementType(new Long(id),"Resource "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
				
				for (CompleteDocuments completeDocuments : CD) {
					boolean docrep=(new Random()).nextBoolean();
					if (docrep)
						{
						
						boolean URL=(new Random()).nextBoolean();
						CompleteResourceElement CTE;
						if (URL)
							CTE=new CompleteResourceElementURL(new Long(id), CX, "URL "+(i*10+j));
						else 
							{
							CompleteFile FF = new CompleteFile(new Long(id), "Path File "+(i*10+j), CC);
							CC.getSectionValues().add(FF);
							id++;
							CTE=new CompleteResourceElementFile(new Long(id), CX, FF);
							}
						id++;
						completeDocuments.getDescription().add(CTE);
						}
				}
				
			}
			  CC.getMetamodelGrammar().add(G1);
		}
		 
		  
		  
		  processCompleteCollection(new CompleteLogAndUpdates(), CC, false, System.getProperty("user.home"));
		  
	    }


	/**
	 *  Retorna el Texto que representa al path.
	 *  @return Texto cadena para el elemento
	 */
	public static String pathFather(CompleteStructure entrada)
	{
		String DataShow;
		if (entrada instanceof CompleteElementType)
			DataShow= ((CompleteElementType) entrada).getName();
		else DataShow= "*";
		
		if (entrada.getFather()!=null)
			return pathFather(entrada.getFather())+"/"+DataShow;
		else return entrada.getCollectionFather().getNombre()+"/"+DataShow;
	}
}
