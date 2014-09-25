/**
 * 
 */
package fdi.ucm.server.exportparser.xls;

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
import fdi.ucm.server.modelComplete.collection.CompleteCollectionLog;
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
public class CollectionXLSI {


	public static String processCompleteCollection(CompleteCollectionLog cL,
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
        
        for (CompleteGrammar row : salvar.getMetamodelGrammar()) {
			processGrammar(libro,row,clave,cL,salvar.getEstructuras(),soloEstructura);
		}
        
        
//        if (!salvar.getName().isEmpty())
//        	 hoja = libro.createSheet(salvar.getName());
//        else hoja = libro.createSheet();
//  
//        
//        ArrayList<CompleteElementType> ListaElementos=generaLista(salvar.getMetamodelGrammar());
//        
//        boolean Errores=false;
//        
//        if (ListaElementos.size()>255)
//        	{
//        	cL.getLogLines().add("Tamaño de estructura demasiado grande para exportar a xls");
//        	Errores=true;
//        	}
//        	
//        if (salvar.getEstructuras().size()+1>65536)
//    	{
//    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls");
//    	Errores=true;
//    	}
//       
//        if (!Errores)
//        {
//        	
//        int row=0;
//        int Column=0;
//        int columnsMax=ListaElementos.size();
//       	
//        
//        for (int i = 0; i < 2; i++) {
//        	Row fila = hoja.createRow(row);
//        	row++;
//        	
//        	for (int j = 0; j < columnsMax; j++) {
//        		
//        		String Value = "";
//            	if (j!=0)
//            		{
//            		CompleteElementType TmpEle = ListaElementos.get(j-1);
//            		Value=pathFather(TmpEle);
//            		}
//            	
//            	if (Value.length()<32767)
//            	{
//            		Cell celda = fila.createCell(j);
//            		
//            		
//            	String Value2 = "";
//            	if (j!=0)
//            		{
//            		Value2=ListaElementos.get(j-1).getClavilenoid().toString();
//            		if (i==0)
//            			clave.put(ListaElementos.get(j-1).getClavilenoid(), Column);
//            		}
//            	
//            	if(i==0){
//            		Column++;
//            		celda.setCellValue(Value);
//            	}else if (i==1){
//            		celda.setCellValue(Value2);
//            	}
//            }
//           }
//		}	
//        	
//        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
//        for(int f=0;f<salvar.getEstructuras().size();f++){
//            /*La clase Row nos permitirá crear las filas*/
//            Row fila = hoja.createRow(row);
//            row++;
//
//            CompleteDocuments Doc=salvar.getEstructuras().get(f);
//            HashMap<Integer, ArrayList<CompleteElement>> ListaClave=new HashMap<Integer, ArrayList<CompleteElement>>();
//            
//            for (CompleteElement elem : Doc.getDescription()) {
//				Integer val=clave.get(elem.getHastype().getClavilenoid());
//				if (val!=null)
//					{
//					ArrayList<CompleteElement> Lis=ListaClave.get(val);
//					if (Lis==null)
//						{
//						Lis=new ArrayList<CompleteElement>();
//						}
//					Lis.add(elem);
//					ListaClave.put(val, Lis);
//					}
//			}
//            
//            
//            
//            /*Cada fila tendrá celdas de datos*/
//            for(int c=0;c<Column;c++){
//            	
//            	String Value = "";
//            	if (c!=0)
//            		{
//            		ArrayList<CompleteElement> temp = ListaClave.get(c);
//            		if (temp!=null)
//            		{
//            		for (CompleteElement completeElement : temp) {
//            			if (!Value.isEmpty())
//            				Value=Value+" "; 
//						if (completeElement instanceof CompleteTextElement)
//							Value=Value+((CompleteTextElement)completeElement).getValue();
//						else if (completeElement instanceof CompleteLinkElement)
//							Value=Value+((CompleteLinkElement)completeElement).getValue().getClavilenoid();
//						else if (completeElement instanceof CompleteResourceElementURL)
//							Value=Value+((CompleteResourceElementURL)completeElement).getValue();
//						else if (completeElement instanceof CompleteResourceElementFile)
//							Value=Value+((CompleteResourceElementFile)completeElement).getValue().getPath();
//							
//					}
//            		}
//            		}
//            	else
//            		{
//            		Value=Doc.getClavilenoid()+"";
//            		}
//            	 
//            	if (Value.length()>=32767)
//            	{
//            		Value="";
//            		cL.getLogLines().add("Temaño de Texto en Valor en elemento " + Value + " excesivo, no debe superar los 32767 caracteres, columna ignorada");
//            	}
//                /*Creamos la celda a partir de la fila actual*/
//                Cell celda = fila.createCell(c);               	
//                	if (c==0)
//                		celda.setCellValue(Value);
//                	else
//                		 celda.setCellValue(Value);
//                    /*Si no es la primera fila establecemos un valor*/
//                	//32.767
//
//                
//            	}
//
//            		
//            		
//            }
//        
//        
//        
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
	
	  private static void processGrammar(Workbook libro, CompleteGrammar grammar,
			HashMap<Long, Integer> clave, CompleteCollectionLog cL, List<CompleteDocuments> list, boolean soloEstructura) {
		  
		   Sheet hoja;
		if (!grammar.getNombre().isEmpty())
	        	 hoja = libro.createSheet(grammar.getNombre());
	        else hoja = libro.createSheet();
	  
	        
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
	       	
	        
	        for (int i = 0; i < 1; i++) {
	        	Row fila = hoja.createRow(row);
	        	row++;
	        	
	        	for (int j = 0; j < columnsMax+2; j++) {
	        		
	        		String Value = "";
	            	if (j==0)
	            		Value="Clavy Document Id";
	            	else 
	            		if (j==1)
	            			Value="Description";
	            		else
	            		{
	            		CompleteElementType TmpEle = ListaElementos.get(j-2);
	            		Value=pathFather(TmpEle);
	            		}
	
	            	
	            	if (Value.length()>=32767)
	            	{
	            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		Value.substring(0, 32766);
	            	}
	            		Cell celda = fila.createCell(j);
	            		
	            		
	            	if (j>1)
	            		{
	            		if (i==0)
	            			clave.put(ListaElementos.get(j-2).getClavilenoid(), Column);
	            		Column++;
	            		}
	            	else
	            	if (j==0)
	            	{
	            		hoja.setColumnWidth(j, 6375);
	            	}
	            	else
	            	if (j==1)
		            {
		            	hoja.setColumnWidth(j, 12750);
		            }
	            	
	            	celda.setCellValue(Value);
	            
	           }
			}	
	        
	        
	        if (!soloEstructura)
	        {
	        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
	        for(int f=0;f<ListaDocumentos.size();f++){
	            /*La clase Row nos permitirá crear las filas*/
	            Row fila = hoja.createRow(row);
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
	            	
	            	String Value = "";
	            	if (c==0)
	            		Value=Long.toString(Doc.getClavilenoid());
	            	else if (c==1)
	            		Value=Doc.getDescriptionText();
	            	else
	            		{
	            		ArrayList<CompleteElement> temp = ListaClave.get(c-2);
	            		if (temp!=null)
	            		{
	            		for (CompleteElement completeElement : temp) {
	            			if (!Value.isEmpty())
	            				Value=Value+" "; 
							if (completeElement instanceof CompleteTextElement)
								Value=Value+((CompleteTextElement)completeElement).getValue();
							else if (completeElement instanceof CompleteLinkElement)
								Value=Value+((CompleteLinkElement)completeElement).getValue().getClavilenoid();
							else if (completeElement instanceof CompleteResourceElementURL)
								Value=Value+((CompleteResourceElementURL)completeElement).getValue();
							else if (completeElement instanceof CompleteResourceElementFile)
								Value=Value+((CompleteResourceElementFile)completeElement).getValue().getPath();
								
						}
	            		}
	            		}
	
	            	 
	            	if (Value.length()>=32767)
	            	{
	            		Value="";
	            		cL.getLogLines().add("Tamaño de Texto en Valor en elemento " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		Value.substring(0, 32766);
	            	}
	                /*Creamos la celda a partir de la fila actual*/
	                Cell celda = fila.createCell(c);               	
	                		 celda.setCellValue(Value);
	                    /*Si no es la primera fila establecemos un valor*/
	                	//32.767

	                
	            	}

	            		
	            		
	            }
	        
	        }
	        
	       
		
	}

	private static ArrayList<CompleteDocuments> generaDocs(
			List<CompleteDocuments> list, CompleteGrammar grammar) {
		ArrayList<CompleteDocuments> ListaDoc=new ArrayList<CompleteDocuments>();
		for (CompleteDocuments completeDocuments : list) {
			if (completeDocuments.getDocument()==grammar)
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

	public static void main(String[] args) throws Exception{
		
		int id=0;
		
		
		
		  CompleteCollection CC=new CompleteCollection("Lou Arreglate", "Arreglate ya!");
		  for (int i = 0; i < 5; i++) {
			  CompleteGrammar G1 = new CompleteGrammar(new Long(id),"Grammar"+i, i+"", CC);
			  
			  ArrayList<CompleteDocuments> CD=new ArrayList<CompleteDocuments>();
			  int docsN=(new Random()).nextInt(5);
			  docsN=docsN+5;
			for (int j = 0; j < docsN; j++) {
				CompleteDocuments CDDD=new CompleteDocuments(new Long(id), CC, G1, "", "");
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
		 
		  
		  
		  processCompleteCollection(new CompleteCollectionLog(), CC, false, System.getProperty("user.home"));
		  
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
