import java.util.*;

public class Pracownik 
	implements Comparable<Pracownik>{
	private String name;
	private int fmno;
	private double totalCost = 0.0;
	private double personalDataCost = 0.0;
	private double personalCallCost = 0.0;
	private int dataTransferred = 0;
	private int dataRoaming = 0;
	List<Integer> numery;
	List<Wolne> urlopy;
	
	public int getDataRoaming() {
		return dataRoaming;
	}

	public void setDataRoaming(int dataRoaming) {
		this.dataRoaming = dataRoaming;
	}
	
	public Double getTotalCost() {
		return totalCost;
	}

	public void setTotalCost(Double totalCost) {
		this.totalCost = totalCost;
	}

	public Double getPersonalDataCost() {
		return personalDataCost;
	}

	public void setPersonalDataCost(Double personalDataCost) {
		this.personalDataCost = personalDataCost;
	}
	public Double getPersonalCallCost() {
		return personalCallCost;
	}

	public void setPersonalCallCost(Double personalCallCost) {
		this.personalCallCost = personalCallCost;
	}

	public int getDataTransferred() {
		return dataTransferred;
	}

	public void setDataTransferred(int dataTransferred) {
		this.dataTransferred = dataTransferred;
	}
	
	public Pracownik(String name, int fmno) {
		super();
		this.name = name;
		this.fmno = fmno;
		numery = new LinkedList<Integer>();
		urlopy = new LinkedList<Wolne>();
	}
	
	public String getName () { return name; }
	public int getFMNO() { return fmno; }
	public List<Integer> getNumery() { return numery; }
	public List<Wolne> getUrlopy() { return urlopy; }
	
	public String toString(){
		String user = "";
		user = name + "(" + fmno + "), numery: ";
		for (Integer i : numery) user += i.toString() + ", ";
		for (Wolne u : urlopy) user += u.start.toString() + " - " + u.end.toString() + ", ";
		return user;
	}
	
	public int compareTo(Pracownik other){
		// funkcja porównuje FMNO
		if(this.fmno == other.fmno) return 0;
		if(this.fmno > other.fmno) return 1;
		return -1;
	}
	
	
	
}
