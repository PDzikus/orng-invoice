import java.util.*;

public class Wolne {
	Date start, end;

	public Wolne (Date start, Date end){
		this.start = start;
		this.end = end;
	}

	public String toString() {
		return start.toString() + " - " + end.toString();
	}
}