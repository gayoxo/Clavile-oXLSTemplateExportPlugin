/**
 * 
 */
package fdi.ucm.server.exportparser.xlstemplate;

import java.io.IOException;
import java.util.ArrayList;

import fdi.ucm.server.modelComplete.ImportExportDataEnum;
import fdi.ucm.server.modelComplete.ImportExportPair;
import fdi.ucm.server.modelComplete.CompleteImportRuntimeException;
import fdi.ucm.server.modelComplete.SaveCollection;
import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.CompleteCollectionLog;

/**
 * @author Joaquin Gayoso-Cabada
 *
 */
public class SaveRemoteCollectionXLSTemp extends SaveCollection {

	
	private String FileO = null;
	private ArrayList<ImportExportPair> Parametros;
	private boolean SoloEstructura;


	public SaveRemoteCollectionXLSTemp() {
		super();
	}
	
	
	/* (non-Javadoc)
	 * @see fdi.ucm.server.SaveCollection#processCollecccion(fdi.ucm.shared.model.collection.Collection)
	 */
	@Override
	public CompleteCollectionLog processCollecccion(CompleteCollection Salvar,
			String PathTemporalFiles) throws CompleteImportRuntimeException {

		CompleteCollectionLog CL=new CompleteCollectionLog();
		try {
			FileO=CollectionXLSTemp.processCompleteCollection(CL,Salvar,SoloEstructura,PathTemporalFiles);
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException("Error en carpeta y escritura del archivo en el servidor");
		}
		return CL;

	}

	/* (non-Javadoc)
	 * @see fdi.ucm.server.SaveCollection#getConfiguracion()
	 */
	@Override
	public ArrayList<ImportExportPair> getConfiguracion() {
		if (Parametros==null)
		{
			ArrayList<ImportExportPair> ListaCampos=new ArrayList<ImportExportPair>();
			ListaCampos.add(new ImportExportPair(ImportExportDataEnum.Boolean, "Exclude Documents Data"));
			Parametros=ListaCampos;
			return ListaCampos;
		}
		else return Parametros;
	}

	/* (non-Javadoc)
	 * @see fdi.ucm.server.SaveCollection#setConfiguracion(java.util.ArrayList)
	 */
	@Override
	public void setConfiguracion(ArrayList<String> DateEntrada) {
		if (DateEntrada!=null)	
		{
			String SoloEstructuraT = DateEntrada.get(0);
			if (SoloEstructuraT.equals(Boolean.toString(true)))
				SoloEstructura=true;
			else 
				SoloEstructura=false;
		}
	}

	/* (non-Javadoc)
	 * @see fdi.ucm.server.SaveCollection#getName()
	 */
	@Override
	public String getName() {
		return "XLS Clavy Template";
	}
	
	
	/**
	 * QUitar caracteres especiales.
	 * @param str texto de entrada.
	 * @return texto sin caracteres especiales.
	 */
	public String RemoveSpecialCharacters(String str) {
		   StringBuilder sb = new StringBuilder();
		   for (int i = 0; i < str.length(); i++) {
			   char c = str.charAt(i);
			   if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_') {
			         sb.append(c);
			      }
		}
		   return sb.toString();
		}



	@Override
	public boolean isFileOutput() {
		return true;
	}


	@Override
	public String FileOutput() {
		return FileO;
	}

}
