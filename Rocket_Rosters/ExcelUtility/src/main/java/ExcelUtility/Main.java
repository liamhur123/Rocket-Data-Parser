package ExcelUtility;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.IOException;


public class Main implements ActionListener{
JFrame frame;
JTextField textfield;
JButton[] buttons = new JButton[2];
JButton RegularRoster,afterDateRegistration;

JPanel panel;
Font myfont = new Font("Ink Free", Font.BOLD,30);
    public static void main(String[] args) {
        Main build = new Main();
    }

 
 Main(){
     frame = new JFrame("Roster Parser"); // creates frame and give it a title
     frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);// when you press red x it just closes it
     frame.setSize(420,550); // sets size of application
     frame.setLayout(null);
     RegularRoster = new JButton("Convert roster");// creates a new button with text saying "convert roster"
     afterDateRegistration = new JButton("Registration Date Search");
     buttons[0] = RegularRoster; // adds button to button array
     buttons[1] = afterDateRegistration;
     // can use for loop to add action listener but for testing just going to do it for one button
     for(int i = 0; i< buttons.length ; i++){
         buttons[i].addActionListener(this);
     }
     afterDateRegistration.setBounds(50,145,300,50);
     RegularRoster.setBounds(50,85,300,50); // sets bound for button

     textfield = new JTextField(); // creates text field
     textfield.setBounds(50,25,300,50); // sets bounds for text field
     textfield.setFont(myfont); // sets the font for the text field
     textfield.setEditable(false); // makes it so user cant type in the text field
     frame.add(textfield); // adds text field to frame
     frame.add(RegularRoster); // adds button to frame
     frame.add(afterDateRegistration);
     frame.setVisible(true); // makes the frame work
} 

    @Override
    public void actionPerformed(ActionEvent e) {
        String PrjDir = System.getProperty("user.dir");
        String excelPath = PrjDir +"/data/active_report.xlsx";
        ExcelUtils excel = new ExcelUtils(excelPath);
        if (e.getSource() == buttons[0]){
            textfield.setText("Names Converted!");
            try {
                excel.WriteClassNamesandTimes();
            } catch (IOException ex) {
                throw new RuntimeException(ex);

            }
        }else {
            textfield.setText("Not supported yet.");
            throw new UnsupportedOperationException("Not supported yet.");}
    }
}