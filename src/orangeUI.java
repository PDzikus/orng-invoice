

import java.awt.GridLayout;

import javax.swing.*;

public class orangeUI {
	JFrame ramkaGlowna = new JFrame();
	
	public orangeUI () {
		ramkaGlowna.setLocation(50, 50);
		ramkaGlowna.setSize(400,200);
		ramkaGlowna.setTitle("Orange - przygotowanie zestawienia");
		ramkaGlowna.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}
	
	public void wyborPlikow (){
		JPanel panel = new JPanel();
		panel.setLayout(new GridLayout(2,2 ,5,5));
		
		JPanel panelSrodkowy = new JPanel();
		panelSrodkowy.setLayout(new GridLayout (2,1,3,3));
		JButton test = new JButton("tester");
		JButton test2 = new JButton("tester2");
		JLabel plikFakturyLabel = new JLabel("Plik faktury: ");
		JLabel plikTelefonyLabel = new JLabel ("Plik z telefonami:");
		JLabel plikUrlopyLabel = new JLabel ("Plik z numerami telefonów");
		
		JTextField plikFaktury = new JTextField("Invoice.xls");
		
		panelSrodkowy.add(plikFakturyLabel);
		panelSrodkowy.add(plikFaktury);
		
		panel.add(panelSrodkowy);
		panel.add(plikUrlopyLabel);
		panel.add(plikTelefonyLabel);
		panel.add(test);
		ramkaGlowna.add(panel);
		ramkaGlowna.setVisible(true);

	}
	
	
}


