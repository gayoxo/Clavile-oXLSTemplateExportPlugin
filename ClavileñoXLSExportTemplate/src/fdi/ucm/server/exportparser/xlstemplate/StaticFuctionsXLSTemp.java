/**
 * 
 */
package fdi.ucm.server.exportparser.xlstemplate;

import java.util.HashSet;
import java.util.List;

import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteStructure;

/**
 * Funcion que implementa las funciones estaticas de la exportacion
 * @author Joaquin Gayoso-Cabada
 *
 */
public class StaticFuctionsXLSTemp {
	
	
	public static boolean isInGrammar(CompleteDocuments iterable_element,
			CompleteGrammar completeGrammar) {
		HashSet<Long> ElemT=new HashSet<Long>();
		for (CompleteElement dd : iterable_element.getDescription()) {
			ElemT.add(dd.getHastype().getClavilenoid());
		}
		
		return isInGrammar(ElemT, completeGrammar.getSons());
		
		
	}



	private static boolean isInGrammar(HashSet<Long> elemT,
			List<CompleteStructure> sons) {
		for (CompleteStructure CSlong1 : sons) {
			if (elemT.contains(CSlong1.getClavilenoid())||isInGrammar(elemT, CSlong1.getSons()))
				return true;
			
		}
		return false;
	}

}
