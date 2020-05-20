//LCD模組函式適用於文字型LCD
//lcd.h
#ifndef _lcd_h_
#define _lcd_h_
#include <util/delay.h>     /* for _delay_ms() */
#include "lcd.h"

/*   =======  硬體接腳的定義  =============  */
#define DBPORT PORTB   
#define CTPORT PORTD   

#define RS_INDEX 0x01 
#define RW_INDEX 0x02
#define E_INDEX  0x08

const char HexData[16] = "0123456789ABCDEF";
const uchar picture1[8] = {0x00,0x0A,0x0A,0x0A,0x00,0x11,0x0E,0x00};

/*   =======  LCD基本顯示驅動  =============     */     
void Init_LCD();                 // LCD 初始化模組
uchar LCD_BUSY(void);            // 0 for Ready , NZ for Busy
void LCD_CMD(char cmd);        // 寫入指令暫存器函式原型宣告
void LCD_DATA(char data1);     // 寫入資料暫存器函式原型宣告
/*   =======  LCD顯示  =============     */     
void DISP_Str(char col,char *str);   //顯示字串
void DISP_Chr(char col,char chrx);  //顯示字元
void DISP_Hex(char addr1,unsigned char v1);
//void RL_Str (char addr1,char *str,char offset);    //字串左旋顯示函式原型宣告
//void RR_Str (char addr1,char *str,char offset);    //字串右移顯示函式原型宣告
void DISP_Int(char addr1,long v1);                   // LCD顯示數字函式原型宣告
void lcd_set_char(uchar address);
 
/*   ==========================     */    
void Init_LCD()               // LCD 初始化函式
{
   _delay_ms(30);     // 等待 LCD 電源開啟完成	

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
	LCD_CMD(0x80);     // 設定DD RAM位址第一行第一列	
	_delay_ms(10);   
	
} 

uchar LCD_BUSY(void)  /* 0 for Ready , NZ for Busy */
{
   uchar busy;
   
   CTPORT = CTPORT |  RW_INDEX;         //RW = 1 
   CTPORT = CTPORT & ~RS_INDEX;         //RS = 0
   CTPORT = CTPORT |  E_INDEX;          //E  = 1
   _delay_us(5);                    // 一個短時間的延遲時序   
   
   //IO設定成輸入
   DDRB  = DDRB & 0xf7;
   PORTB = PORTB | 0x08; 
   
   _delay_us(5);                    // 一個短時間的延遲時序   
   busy = PINB & 0x08;
   
   CTPORT = CTPORT & ~E_INDEX;          //E  = 0
  
   //IO設定成輸出   
   DDRB  = DDRB | 0x08;   
     
   return (busy);
  
}

void LCD_CMD(char cmd)        // 寫入指令暫存器函式
{
   while (LCD_BUSY());
   
	CTPORT = CTPORT & ~RS_INDEX;                //RS = 0
   _delay_us(5);                    // 一個短時間的延遲時序		
	CTPORT = CTPORT & ~RW_INDEX;                //RW = 0	
   _delay_us(5);                    // 一個短時間的延遲時序	

	DBPORT = (DBPORT & 0xF0)|((cmd>>4) & 0x0F); //寫入指令暫存器  send higher nibble		
   _delay_us(5);                    // 一個短時間的延遲時序		
	CTPORT = CTPORT |  E_INDEX;                 //E  = 1		
	_delay_us(5);                    // 一個短時間的延遲時序	
	CTPORT = CTPORT & ~E_INDEX;                 //E  = 0		
   _delay_us(5);                    // 一個短時間的延遲時序		

	DBPORT = (DBPORT & 0xF0)|(cmd & 0x0F);      //寫入指令暫存器  send lower nibble
	_delay_us(5);                    // 一個短時間的延遲時序	
	CTPORT = CTPORT |  E_INDEX;                 //E  = 1		
	_delay_us(5);                    // 一個短時間的延遲時序	
	CTPORT = CTPORT & ~E_INDEX;                 //E  = 0	
	_delay_us(5);                    // 一個短時間的延遲時序			
}

void LCD_DATA(char data1)     // 寫入資料暫存器函式
{
   while (LCD_BUSY());
   
	CTPORT = CTPORT |  RS_INDEX;         //RS = 1
	_delay_us(5);                    // 一個短時間的延遲時序		
	CTPORT = CTPORT & ~RW_INDEX;         //RW = 0	
	_delay_us(5);                    // 一個短時間的延遲時序		
	
	DBPORT = (DBPORT & 0xF0)|((data1>>4) & 0x0F); //寫入指令暫存器  send higher nibble		
   _delay_us(5);                    // 一個短時間的延遲時序		
	CTPORT = CTPORT |  E_INDEX;                 //E  = 1		
	_delay_us(5);                    // 一個短時間的延遲時序		
	CTPORT = CTPORT & ~E_INDEX;                 //E  = 0		
   _delay_us(5);                    // 一個短時間的延遲時序		

	DBPORT = (DBPORT & 0xF0)|(data1 & 0x0F);      //寫入指令暫存器  send lower nibble
   _delay_us(5);                    // 一個短時間的延遲時序	     	
	CTPORT = CTPORT |  E_INDEX;                 //E  = 1		
	_delay_us(5);                    // 一個短時間的延遲時序	
	CTPORT = CTPORT & ~E_INDEX;                 //E  = 0	
   _delay_us(5);                    // 一個短時間的延遲時序		


}

void DISP_Str(char addr1,char *str)         // 在LCD指定位置顯示字串函式
{
	LCD_CMD(addr1);         // 設定DD RAM位址第一行第一列

	while(*str !=0)
	    LCD_DATA(*str++);	// 呼叫顯示字串函式
}

void DISP_Chr(char col,char chrx)           // 在LCD指定位置顯示字元函式 
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
	LCD_CMD(addr1) ;         // 設定DD RAM位址
	while(*str !=0)	
            LCD_DATA(*str++); 	// 呼叫顯示字串函式	     
  	for (i=0;i<offset;i++)
	    LCD_DATA(start[i]);		// 呼叫顯示字串函式	          
}

//字串右旋顯示函式
void RR_Str (char addr1,char *str,char offset)
{
	char i;
	char *start=str;

	LCD_CMD(addr1) ;         // 設定DD RAM位址為0
	str+=20-offset;

	while(*str !=0)
	    LCD_DATA(*str++);	// 呼叫顯示字串函式	     
        for (i=0;i<20-offset;i++)
	    LCD_DATA(start[i]);	// 呼叫顯示字串函式	     
}
*/

void DISP_Int(char addr1,long v1)           // 在LCD指定位置顯示字元函式  
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

void DISP_Hex(char addr1,unsigned char v1)  // 在LCD指定位置顯示字元函式
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
