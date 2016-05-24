
import java.util.*;


public class Rozmowa 
	implements java.lang.Comparable<Rozmowa> {
	String nrTelefonu;
	Date data;
	String kraj;
	String typ;
	String opis;
	String wybranyNumer;
	String czas;
	double koszt;
	
	
	public Rozmowa(String nrTelefonu, Date data, String kraj, String typ, String opis, String wybranyNumer, String czas,
			double koszt) {
		this.nrTelefonu = nrTelefonu;
		this.data = data;
		this.kraj = kraj;
		this.typ = typ;
		this.opis = opis;
		this.wybranyNumer = wybranyNumer;
		this.czas = czas;
		this.koszt = koszt;
	}
	
	public String getNrTelefonu (){
		return nrTelefonu;
	}
	
	public int compareTo(Rozmowa r){
		return this.data.compareTo(r.data);
	}
	
	public String toString(){
		return nrTelefonu + ", " + data.toString()+ ", kraj:" + kraj + ", " + opis + ", " + wybranyNumer + ", " + czas + ", " + koszt;
	}
	
}
