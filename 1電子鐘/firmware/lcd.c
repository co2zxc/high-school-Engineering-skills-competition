//LCD�Ҳը禡�A�Ω��r��LCD
//lcd.h
#ifndef _lcd_h_
#define _lcd_h_
#include <util/delay.h>     /* for _delay_ms() */
#include "lcd.h"

/*   =======  �w�鱵�}���w�q  =============  */
#define DBPORT PORTB   
#define CTPORT PORTD   

#define RS_INDEX 0x01 
#define RW_INDEX 0x02
#define E_INDEX  0x08

const char HexData[16] = "0123456789ABCDEF";
const uchar picture1[8] = {0x00,0x0A,0x0A,0x0A,0x00,0x11,0x0E,0x00};

/*   =======  LCD������X��  =============     */     
void Init_LCD();                 // LCD ��l�ƼҲ�
uchar LCD_BUSY(void);            // 0 for Ready , NZ for Busy
void LCD_CMD(char cmd);        // �g�J���O�Ȧs���禡�쫬�ŧi
void LCD_DATA(char data1);     // �g�J��ƼȦs���禡�쫬�ŧi
/*   =======  LCD���  =============     */     
void DISP_Str(char col,char *str);   //��ܦr��
void DISP_Chr(char col,char chrx);  //��ܦr��
void DISP_Hex(char addr1,unsigned char v1);
//void RL_Str (char addr1,char *str,char offset);    //�r�ꥪ����ܨ禡�쫬�ŧi
//void RR_Str (char addr1,char *str,char offset);    //�r��k����ܨ禡�쫬�ŧi
void DISP_Int(char addr1,long v1);                   // LCD��ܼƦr�禡�쫬�ŧi
void lcd_set_char(uchar address);
 
/*   ==========================     */    
void Init_LCD()               // LCD ��l�ƨ禡
{
   _delay_ms(30);     // ���� LCD �q���}�ҧ���	

	LCD_CMD(0x33);		
	LCD_CMD(0x32);			
   
	LCD_CMD(0x28);		
	_delay_ms(1);   
//	LCD_CMD(0x0E);
	LCD_CMD(0x0C);
	_delay_ms(1);   
	LCD_CMD(0x01);			
	_delay_ms(1);   
	LCD_CMD(0x06);			
	_delay_ms(10);   
	LCD_CMD(0x80);     // �]�wDD RAM��}�Ĥ@��Ĥ@�C	
	_delay_ms(10);   
	
} 

uchar LCD_BUSY(void)  /* 0 for Ready , NZ for Busy */
{
   uchar busy;
   
   CTPORT = CTPORT |  RW_INDEX;         //RW = 1 
   CTPORT = CTPORT & ~RS_INDEX;         //RS = 0
   CTPORT = CTPORT |  E_INDEX;          //E  = 1
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�   
   
   //IO�]�w����J
   DDRB  = DDRB & 0xf7;
   PORTB = PORTB | 0x08; 
   
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�   
   busy = PINB & 0x08;
   
   CTPORT = CTPORT & ~E_INDEX;          //E  = 0
  
   //IO�]�w����X   
   DDRB  = DDRB | 0x08;   
     
   return (busy);
  
}

void LCD_CMD(char cmd)        // �g�J���O�Ȧs���禡
{
   while (LCD_BUSY());
   
	CTPORT = CTPORT & ~RS_INDEX;                //RS = 0
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�		
	CTPORT = CTPORT & ~RW_INDEX;                //RW = 0	
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�	

	DBPORT = (DBPORT & 0xF0)|((cmd>>4) & 0x0F); //�g�J���O�Ȧs��  send higher nibble		
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�		
	CTPORT = CTPORT |  E_INDEX;                 //E  = 1		
	_delay_us(5);                    // �@�ӵu�ɶ�������ɧ�	
	CTPORT = CTPORT & ~E_INDEX;                 //E  = 0		
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�		

	DBPORT = (DBPORT & 0xF0)|(cmd & 0x0F);      //�g�J���O�Ȧs��  send lower nibble
	_delay_us(5);                    // �@�ӵu�ɶ�������ɧ�	
	CTPORT = CTPORT |  E_INDEX;                 //E  = 1		
	_delay_us(5);                    // �@�ӵu�ɶ�������ɧ�	
	CTPORT = CTPORT & ~E_INDEX;                 //E  = 0	
	_delay_us(5);                    // �@�ӵu�ɶ�������ɧ�			
}

void LCD_DATA(char data1)     // �g�J��ƼȦs���禡
{
   while (LCD_BUSY());
   
	CTPORT = CTPORT |  RS_INDEX;         //RS = 1
	_delay_us(5);                    // �@�ӵu�ɶ�������ɧ�		
	CTPORT = CTPORT & ~RW_INDEX;         //RW = 0	
	_delay_us(5);                    // �@�ӵu�ɶ�������ɧ�		
	
	DBPORT = (DBPORT & 0xF0)|((data1>>4) & 0x0F); //�g�J���O�Ȧs��  send higher nibble		
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�		
	CTPORT = CTPORT |  E_INDEX;                 //E  = 1		
	_delay_us(5);                    // �@�ӵu�ɶ�������ɧ�		
	CTPORT = CTPORT & ~E_INDEX;                 //E  = 0		
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�		

	DBPORT = (DBPORT & 0xF0)|(data1 & 0x0F);      //�g�J���O�Ȧs��  send lower nibble
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�	     	
	CTPORT = CTPORT |  E_INDEX;                 //E  = 1		
	_delay_us(5);                    // �@�ӵu�ɶ�������ɧ�	
	CTPORT = CTPORT & ~E_INDEX;                 //E  = 0	
   _delay_us(5);                    // �@�ӵu�ɶ�������ɧ�		


}

void DISP_Str(char addr1,char *str)         // �bLCD���w��m��ܦr��禡
{
	LCD_CMD(addr1);         // �]�wDD RAM��}�Ĥ@��Ĥ@�C

	while(*str !=0)
	    LCD_DATA(*str++);	// �I�s��ܦr��禡
}

void DISP_Chr(char col,char chrx)           // �bLCD���w��m��ܦr���禡 
{
	LCD_CMD(col);         
	LCD_DATA(chrx);	                    
}

/*
void RL_Str(char addr1,char *str,char offset)
{
	char i;
	char *start=str;

	str+=offset;
	LCD_CMD(addr1) ;         // �]�wDD RAM��}
	while(*str !=0)	
            LCD_DATA(*str++); 	// �I�s��ܦr��禡	     
  	for (i=0;i<offset;i++)
	    LCD_DATA(start[i]);		// �I�s��ܦr��禡	          
}

//�r��k����ܨ禡
void RR_Str (char addr1,char *str,char offset)
{
	char i;
	char *start=str;

	LCD_CMD(addr1) ;         // �]�wDD RAM��}��0
	str+=20-offset;

	while(*str !=0)
	    LCD_DATA(*str++);	// �I�s��ܦr��禡	     
        for (i=0;i<20-offset;i++)
	    LCD_DATA(start[i]);	// �I�s��ܦr��禡	     
}
*/

void DISP_Int(char addr1,long v1)           // �bLCD���w��m��ܦr���禡  
{
	char i;
	for (i=0;i<7;i++){
		LCD_CMD(addr1-i+7);
		LCD_DATA((v1%10)+0x30);
		v1/=10;
		if(v1==0)
		break;
	}
}

void DISP_Hex(char addr1,unsigned char v1)  // �bLCD���w��m��ܦr���禡
{
   unsigned char i,j;
   i = v1 / 16 ;
   j = v1 % 16 ;
   DISP_Chr(addr1,'0');
   DISP_Chr(addr1 + 1,'x');   
   DISP_Chr(addr1 + 2 ,HexData[i]);
   DISP_Chr(addr1 + 3 ,HexData[j]);
}

void lcd_set_char(uchar address)            // address = 0 -7
{
   uchar i;

   LCD_CMD( 0x40 | (address*8) );           //set address   
   for(i=0;i<8;i++)         	
	    LCD_DATA(picture1[i]);	     
	
}

#endif
