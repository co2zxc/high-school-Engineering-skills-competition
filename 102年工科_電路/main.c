#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* for sei() */
#include <util/delay.h>     /* for _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* required by usbdrv.h */
#include "usbdrv.h"
#include "oddebug.h"        /* This is also an example for using debug macros */

#include "Time_Number.h"
#include "Math_Number.h"
#include "Moving_Logo.h"

static uchar VB_mod;
static uchar keyin;
static uchar LED[4][8];
static horse_first;
static uchar time;
static uchar horse_count;
static uchar horse_LR;
	
static void outputEnable(uchar b)
{
   	//8*8 LED Enable
	PORTD = (PORTD & 0xFD); //0xFD 11111101
	PORTC = (PORTC & 0xC0) | (b & 0x3F); //0xC0 11000000  //0x3F 00111111
	PORTD = (PORTD & 0xFD) | ((b >> 5) & 0x02); //0xFD 11111101  //0x02 00000010
}

static void outputLED(uchar b)
{
	//8 bit LED
	PORTB = (PORTB & 0xC0) | (b & 0x3F); //0xC0 11000000  //0x3F 00111111
	PORTD = (PORTD & 0x3F) | (b & 0xC0); //0x3F 00111111  //0xC0 11000000
}
// Clock
static void outputControl()
{
	PORTD = PORTD | 0x20; //0x20 00100000
	PORTD = PORTD & 0xDF; //0xDF 11011111
}

static uchar inputKeys_1()
{
	uchar b;
	b = PIND & 0x01;
	return b;
}

static uchar inputKeys_2()
{
	uchar b;
	b = (PIND >> 3) & 0x01;
	return b;
}

static void Matrix(uchar Row, uchar Data)
{
	//8*8 LED
		uchar i; 
		for(i = 0 ; i < 8 ; i++)
		{
			PORTB =(PORTB & 0xC0) | (LED[Data][i] & 0x3f); //0xC0 11000000 //0x3f 00111111
			PORTD =(PORTD & 0x3F) | (LED[Data][i] & 0xC0); //0x3F 00111111 //0xC0 11000000  
			outputEnable(Row + i);
			_delay_ms(0.01);
			outputEnable(0x00);
		}	
}

static void Matrix_LED(uchar n, uchar Row,uchar Data)  // 第n顆  第Row行  資料Data
{
	LED[n][Row] = Data;
}

static void horse_left(uchar data)
{
	//8*8 horse to left
	LED[0][0] = LED[0][1];
	LED[0][1] = LED[0][2];	
	LED[0][2] = LED[0][3];
	LED[0][3] = LED[0][4];
	LED[0][4] = LED[0][5];
	LED[0][5] = LED[0][6];
	LED[0][6] = LED[0][7];
	LED[0][7] = LED[1][0];

	LED[1][0] = LED[1][1];
	LED[1][1] = LED[1][2];	
	LED[1][2] = LED[1][3];
	LED[1][3] = LED[1][4];
	LED[1][4] = LED[1][5];
	LED[1][5] = LED[1][6];
	LED[1][6] = LED[1][7];
	LED[1][7] = LED[2][0];

	LED[2][0] = LED[2][1];
	LED[2][1] = LED[2][2];	
	LED[2][2] = LED[2][3];
	LED[2][3] = LED[2][4];
	LED[2][4] = LED[2][5];
	LED[2][5] = LED[2][6];
	LED[2][6] = LED[2][7];
	LED[2][7] = LED[3][0];
		
	LED[3][0] = LED[3][1];
	LED[3][1] = LED[3][2];	
	LED[3][2] = LED[3][3];
	LED[3][3] = LED[3][4];
	LED[3][4] = LED[3][5];
	LED[3][5] = LED[3][6];
	LED[3][6] = LED[3][7];
	LED[3][7] = data;	
}

static void horse_right(uchar data)
{
	LED[3][7] = LED[3][6];
	LED[3][6] = LED[3][5];	
	LED[3][5] = LED[3][4];
	LED[3][4] = LED[3][3];
	LED[3][3] = LED[3][2];
	LED[3][2] = LED[3][1];
	LED[3][1] = LED[3][0];
	LED[3][0] = LED[2][7];
			
	LED[2][7] = LED[2][6];
	LED[2][6] = LED[2][5];	
	LED[2][5] = LED[2][4];
	LED[2][4] = LED[2][3];
	LED[2][3] = LED[2][2];
	LED[2][2] = LED[2][1];
	LED[2][1] = LED[2][0];
	LED[2][0] = LED[1][7];
			
	LED[1][7] = LED[1][6];
	LED[1][6] = LED[1][5];	
	LED[1][5] = LED[1][4];
	LED[1][4] = LED[1][3];
	LED[1][3] = LED[1][2];
	LED[1][2] = LED[1][1];
	LED[1][1] = LED[1][0];
	LED[1][0] = LED[0][7];
			
	LED[0][7] = LED[0][6];
	LED[0][6] = LED[0][5];	
	LED[0][5] = LED[0][4];
	LED[0][4] = LED[0][3];
	LED[0][3] = LED[0][2];
	LED[0][2] = LED[0][1];
	LED[0][1] = LED[0][0];
	LED[0][0] = data ;

}

static void Moving_Logo_fist_set()
{
	Matrix_LED(0,0,moving_logo[0]);	
	Matrix_LED(0,1,moving_logo[1]);	
	Matrix_LED(0,2,moving_logo[2]);	
	Matrix_LED(0,3,moving_logo[3]);	
	Matrix_LED(0,4,moving_logo[4]);	
	Matrix_LED(0,5,moving_logo[5]);	
	Matrix_LED(0,6,moving_logo[6]);	
	Matrix_LED(0,7,moving_logo[7]);	
		
	Matrix_LED(1,0,moving_logo[8]);	
	Matrix_LED(1,1,moving_logo[9]);	
	Matrix_LED(1,2,moving_logo[10]);	
	Matrix_LED(1,3,moving_logo[11]);	
	Matrix_LED(1,4,moving_logo[12]);	
	Matrix_LED(1,5,moving_logo[13]);	
	Matrix_LED(1,6,moving_logo[14]);	
	Matrix_LED(1,7,moving_logo[15]);		
		
	Matrix_LED(2,0,moving_logo[16]);	
	Matrix_LED(2,1,moving_logo[17]);	
	Matrix_LED(2,2,moving_logo[18]);	
	Matrix_LED(2,3,moving_logo[19]);	
	Matrix_LED(2,4,moving_logo[20]);	
	Matrix_LED(2,5,moving_logo[21]);	
	Matrix_LED(2,6,moving_logo[22]);	
	Matrix_LED(2,7,moving_logo[23]);	
		
	Matrix_LED(3,0,moving_logo[24]);	
	Matrix_LED(3,1,moving_logo[25]);	
	Matrix_LED(3,2,moving_logo[26]);	
	Matrix_LED(3,3,moving_logo[27]);	
	Matrix_LED(3,4,moving_logo[28]);	
	Matrix_LED(3,5,moving_logo[29]);	
	Matrix_LED(3,6,moving_logo[30]);	
	Matrix_LED(3,7,moving_logo[31]);	
			
	horse_count = 32;	
	horse_first = 1;
	horse_LR = 0;
	time = 0;
			
}

static void Moving_Logo_set()
{
	
	time = time + 1 ;				
	if (time == 200)
	{
		switch(horse_LR)
		{
			case 0:
				horse_left(moving_logo[horse_count]);
			break;	
			
			case 1:
				horse_right(moving_logo[horse_count]);
			break;	
		}	
		time =0;
		horse_count = horse_count + 1 ;
		if (horse_count == 64)
		{
			horse_LR = 1 ; //right
		}
		else if (horse_count == 128)
		{
			horse_count = 0 ;
			horse_LR = 0 ; //left
		}
		
	}
}

static void VB_0()
{
	uchar B0;
	B0 = PIND & 0x01;
	
	uchar B1;
	B1 = (PIND >> 2) & 0x02;
	
	uchar B1_B0;
	B1_B0 = B1 + B0;	
	
	if (B1_B0 != 0 ) {
		horse_first = 0;
	}
	
	switch(B1_B0)
	{	
	
	case 2:  //OFF ON
		
		//
		Matrix_LED(0,0,0x0);
		Matrix_LED(0,1,0x0);
		Matrix_LED(0,2,0x0);
		Matrix_LED(0,3,0x0);
		Matrix_LED(0,4,0x0);
		Matrix_LED(0,5,0x0);
		Matrix_LED(0,6,0x0);
		Matrix_LED(0,7,0x0);
		
		//
		Matrix_LED(1,0,0x0);
		Matrix_LED(1,1,0x0);
		Matrix_LED(1,2,0x0);
		Matrix_LED(1,3,0x0);
		Matrix_LED(1,4,0x0);
		Matrix_LED(1,5,0x0);
		Matrix_LED(1,6,0x0);	
		
		//:: (10)
		Matrix_LED(1,7,math_number[10][0]);
		Matrix_LED(2,0,math_number[10][0]);
		
		//
		Matrix_LED(2,1,0x0);
		Matrix_LED(2,2,0x0);
		Matrix_LED(2,3,math_number[4][0]);
		Matrix_LED(2,4,math_number[4][1]);
		Matrix_LED(2,5,math_number[4][2]);
		Matrix_LED(2,6,math_number[4][3]);
		Matrix_LED(2,7,math_number[4][4]);
		
		//9
		Matrix_LED(3,0,0x0);
		Matrix_LED(3,1,0x0);
		Matrix_LED(3,2,math_number[7][0]);
		Matrix_LED(3,3,math_number[7][1]);
		Matrix_LED(3,4,math_number[7][2]);
		Matrix_LED(3,5,math_number[7][3]);
		Matrix_LED(3,6,math_number[7][4]);
		Matrix_LED(3,7,0x0);
		
	break;
	
	case 1: //ON OFF
	//
		Matrix_LED(0,0,0x0);
		Matrix_LED(0,1,0x0);
		Matrix_LED(0,2,0x0);
		Matrix_LED(0,3,0x0);
		Matrix_LED(0,4,0x0);
		Matrix_LED(0,5,0x0);
		Matrix_LED(0,6,0x0);
		Matrix_LED(0,7,0x0);
		
		//
		Matrix_LED(1,0,0x0);
		Matrix_LED(1,1,0x0);
		Matrix_LED(1,2,0x0);
		Matrix_LED(1,3,0x0);
		Matrix_LED(1,4,0x0);
		Matrix_LED(1,5,0x0);
		Matrix_LED(1,6,0x0);
		Matrix_LED(1,7,0x0);
		
		//
		Matrix_LED(2,0,0x0);
		Matrix_LED(2,1,0x0);
		Matrix_LED(2,2,0x0);
		Matrix_LED(2,3,math_number[4][0]);
		Matrix_LED(2,4,math_number[4][1]);
		Matrix_LED(2,5,math_number[4][2]);
		Matrix_LED(2,6,math_number[4][3]);
		Matrix_LED(2,7,math_number[4][4]);
		
		//9
		Matrix_LED(3,0,0x0);
		Matrix_LED(3,1,0x0);
		Matrix_LED(3,2,math_number[9][0]);
		Matrix_LED(3,3,math_number[9][1]);
		Matrix_LED(3,4,math_number[9][2]);
		Matrix_LED(3,5,math_number[9][3]);
		Matrix_LED(3,6,math_number[9][4]);
		Matrix_LED(3,7,0x0);
		
	break;
	
	case 0: //ON ON
	
		switch (horse_first)
		{
		case 0:
			Moving_Logo_fist_set();
		break ;	
		
		case 1:
			Moving_Logo_set();
		break;	
		}
			
	break;
		
	case 3: //OFF OFF
		
	//2
		Matrix_LED(0,0,0x0);
		Matrix_LED(0,1,math_number[2][0]);
		Matrix_LED(0,2,math_number[2][1]);
		Matrix_LED(0,3,math_number[2][2]);
		Matrix_LED(0,4,math_number[2][3]);
		Matrix_LED(0,5,math_number[2][4]);
		Matrix_LED(0,6,0x0);
		Matrix_LED(0,7,0x0);
		
		//3
		Matrix_LED(1,0,math_number[3][0]);
		Matrix_LED(1,1,math_number[3][1]);
		Matrix_LED(1,2,math_number[3][2]);
		Matrix_LED(1,3,math_number[3][3]);
		Matrix_LED(1,4,math_number[3][4]);
		Matrix_LED(1,5,0x0);
		Matrix_LED(1,6,0x0);	
		
		//:: (10)
		Matrix_LED(1,7,math_number[10][0]);
		Matrix_LED(2,0,math_number[10][0]);

		//5
		Matrix_LED(2,1,0x0);
		Matrix_LED(2,2,0x0);
		Matrix_LED(2,3,math_number[5][0]);
		Matrix_LED(2,4,math_number[5][1]);
		Matrix_LED(2,5,math_number[5][2]);
		Matrix_LED(2,6,math_number[5][3]);
		Matrix_LED(2,7,math_number[5][4]);
		
		//9
		Matrix_LED(3,0,0x0);
		Matrix_LED(3,1,0x0);
		Matrix_LED(3,2,math_number[9][0]);
		Matrix_LED(3,3,math_number[9][1]);
		Matrix_LED(3,4,math_number[9][2]);
		Matrix_LED(3,5,math_number[9][3]);
		Matrix_LED(3,6,math_number[9][4]);
		Matrix_LED(3,7,0x0);
		
	break;
		
	}	

}

static void VB_1()
{
	uchar B0;
	B0 = PIND & 0x01;
	
	uchar B1;
	B1 = (PIND >> 2) & 0x02;
	
	uchar B1_B0;
	B1_B0 = B1 + B0;	
	
	if (B1_B0 != 0 ) 	
	{	
		horse_first = 0;
	}	
	else
	{
		switch (horse_first)
		{
		case 0:
			Moving_Logo_fist_set();
		break ;	
			
		case 1:
			Moving_Logo_set();
		break;	
		}
	}
}

static void VB_2()
{
	uchar B0;
	B0 = PIND & 0x01;
	
	uchar B1;
	B1 = (PIND >> 2) & 0x02;
	
	uchar B1_B0;
	B1_B0 = B1 + B0;	
	
	if (B1_B0 != 0 ) {
		time = 0;
		horse_first = 0;
		horse_count = 0;
		horse_LR = 0;
	}
	
	switch(B1_B0)
	{	
	
	case 2:  //OFF ON
		
		//
		Matrix_LED(0,0,0x0);
		Matrix_LED(0,1,0x0);
		Matrix_LED(0,2,0x0);
		Matrix_LED(0,3,0x0);
		Matrix_LED(0,4,0x0);
		Matrix_LED(0,5,0x0);
		Matrix_LED(0,6,0x0);
		Matrix_LED(0,7,0x0);
		
		//
		Matrix_LED(1,0,0x0);
		Matrix_LED(1,1,0x0);
		Matrix_LED(1,2,0x0);
		Matrix_LED(1,3,0x0);
		Matrix_LED(1,4,0x0);
		Matrix_LED(1,5,0x0);
		Matrix_LED(1,6,0x0);	
		
		//:: (10)
		Matrix_LED(1,7,math_number[10][0]);
		Matrix_LED(2,0,math_number[10][0]);
		
		//
		Matrix_LED(2,1,0x0);
		Matrix_LED(2,2,0x0);
		Matrix_LED(2,3,math_number[0][0]);
		Matrix_LED(2,4,math_number[0][1]);
		Matrix_LED(2,5,math_number[0][2]);
		Matrix_LED(2,6,math_number[0][3]);
		Matrix_LED(2,7,math_number[0][4]);
		
		//9
		Matrix_LED(3,0,0x0);
		Matrix_LED(3,1,0x0);
		Matrix_LED(3,2,math_number[0][0]);
		Matrix_LED(3,3,math_number[0][1]);
		Matrix_LED(3,4,math_number[0][2]);
		Matrix_LED(3,5,math_number[0][3]);
		Matrix_LED(3,6,math_number[0][4]);
		Matrix_LED(3,7,0x0);
		
	break;
	
	case 1: //ON OFF
	//
		Matrix_LED(0,0,0x0);
		Matrix_LED(0,1,0x0);
		Matrix_LED(0,2,0x0);
		Matrix_LED(0,3,0x0);
		Matrix_LED(0,4,0x0);
		Matrix_LED(0,5,0x0);
		Matrix_LED(0,6,0x0);
		Matrix_LED(0,7,0x0);
		
		//
		Matrix_LED(1,0,0x0);
		Matrix_LED(1,1,0x0);
		Matrix_LED(1,2,0x0);
		Matrix_LED(1,3,0x0);
		Matrix_LED(1,4,0x0);
		Matrix_LED(1,5,0x0);
		Matrix_LED(1,6,0x0);
		Matrix_LED(1,7,0x0);
		
		//
		Matrix_LED(2,0,0x0);
		Matrix_LED(2,1,0x0);
		Matrix_LED(2,2,0x0);
		Matrix_LED(2,3,math_number[0][0]);
		Matrix_LED(2,4,math_number[0][1]);
		Matrix_LED(2,5,math_number[0][2]);
		Matrix_LED(2,6,math_number[0][3]);
		Matrix_LED(2,7,math_number[0][4]);
		
		//9
		Matrix_LED(3,0,0x0);
		Matrix_LED(3,1,0x0);
		Matrix_LED(3,2,math_number[0][0]);
		Matrix_LED(3,3,math_number[0][1]);
		Matrix_LED(3,4,math_number[0][2]);
		Matrix_LED(3,5,math_number[0][3]);
		Matrix_LED(3,6,math_number[0][4]);
		Matrix_LED(3,7,0x0);
		
	break;
	
	case 3: //OFF OFF
		
	//2
		Matrix_LED(0,0,0x0);
		Matrix_LED(0,1,math_number[0][0]);
		Matrix_LED(0,2,math_number[0][1]);
		Matrix_LED(0,3,math_number[0][2]);
		Matrix_LED(0,4,math_number[0][3]);
		Matrix_LED(0,5,math_number[0][4]);
		Matrix_LED(0,6,0x0);
		Matrix_LED(0,7,0x0);
		
		//3
		Matrix_LED(1,0,math_number[0][0]);
		Matrix_LED(1,1,math_number[0][1]);
		Matrix_LED(1,2,math_number[0][2]);
		Matrix_LED(1,3,math_number[0][3]);
		Matrix_LED(1,4,math_number[0][4]);
		Matrix_LED(1,5,0x0);
		Matrix_LED(1,6,0x0);	
		
		//:: (10)
		Matrix_LED(1,7,math_number[10][0]);
		Matrix_LED(2,0,math_number[10][0]);

		//5
		Matrix_LED(2,1,0x0);
		Matrix_LED(2,2,0x0);
		Matrix_LED(2,3,math_number[0][0]);
		Matrix_LED(2,4,math_number[0][1]);
		Matrix_LED(2,5,math_number[0][2]);
		Matrix_LED(2,6,math_number[0][3]);
		Matrix_LED(2,7,math_number[0][4]);
		
		//9
		Matrix_LED(3,0,0x0);
		Matrix_LED(3,1,0x0);
		Matrix_LED(3,2,math_number[0][0]);
		Matrix_LED(3,3,math_number[0][1]);
		Matrix_LED(3,4,math_number[0][2]);
		Matrix_LED(3,5,math_number[0][3]);
		Matrix_LED(3,6,math_number[0][4]);
		Matrix_LED(3,7,0x0);
		
	break;
		
	}	
}




/* ------------------------------------------------------------------------- */
/* ----------------------------- USB interface ----------------------------- */
/* ------------------------------------------------------------------------- */

PROGMEM const char usbHidReportDescriptor[22] = {    /* USB report descriptor */
    0x06, 0x00, 0xff,              // USAGE_PAGE (Generic Desktop)
    0x09, 0x01,                    // USAGE (Vendor Usage 1)
    0xa1, 0x01,                    // COLLECTION (Application)
    0x15, 0x00,                    // LOGICAL_MINIMUM (0)
    0x26, 0xff, 0x00,              // LOGICAL_MAXIMUM (255)
    0x75, 0x08,                    // REPORT_SIZE (8)
    0x95, 0x80,                    // REPORT_COUNT (128)
    0x09, 0x00,                    // USAGE (Undefined)
    0xb2, 0x02, 0x01,              // FEATURE (Data,Var,Abs,Buf)
    0xc0                           // END_COLLECTION
};

/* Since we define only one feature report, we don't use report-IDs (which
 * would be the first byte of the report). The entire report consists of 128
 * opaque data bytes.
 */

/* The following variables store the status of the current data transfer */
static uchar    currentAddress;
static uchar    bytesRemaining;

/* ------------------------------------------------------------------------- */

/* usbFunctionRead() is called when the host requests a chunk of data from
 * the device. For more information see the documentation in usbdrv/usbdrv.h.
 */
uchar   usbFunctionRead(uchar *data, uchar len)
{
//ReadData
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		*(data) = inputKeys_1();	// Key value
		*(data+1) = inputKeys_2();		
		*(data+2) = 0;		// Reserve (zero)
		*(data+3) = 0;		// Reserve (zero)
		*(data+4) = 0;		// Reserve (zero)
		*(data+5) = 0;		// Reserve (zero)
		*(data+6) = 0;		// Reserve (zero)
		*(data+7) = 0;		// Reserve (zero)
	}
    currentAddress += len;
    bytesRemaining -= len;
    return len;
}

/* usbFunctionWrite() is called when the host sends a chunk of data to the
 * device. For more information see the documentation in usbdrv/usbdrv.h.
 */
uchar   usbFunctionWrite(uchar *data, uchar len)
{
	if(bytesRemaining == 0)
        return 1;               /* end of transfer */
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){

	// OutDataEightByte
	uchar i;
	uchar j;

	switch(*(data))
	{
		case 0x20:
		//8*8 LED
			LED[*(data + 1)][*(data + 2)] = *(data + 3);
		break;
			
		case 0x21:
		//8 bit LED
			outputLED(data[1]); // outputLED(*(data+1))
			outputControl();			// Clock
		break;	
			
		case 0x22:
		//8*8 horse to left
			horse_left(*(data + 1));			
		   break;	

		case 0x23:			
		//8*8 horse to right
			horse_right(*(data + 1));	
		break;	
			
		case 0x24:
		//time HH:MM:SS
			LED[0][2] = time_number[*(data + 1)][0];
			LED[0][3] = time_number[*(data + 1)][1];
			LED[0][4] = time_number[*(data + 1)][2];
			
			LED[0][6] = time_number[*(data + 2)][0];
			LED[0][7] = time_number[*(data + 2)][1];
			LED[1][0] = time_number[*(data + 2)][2];
			
			LED[1][4] = time_number[*(data + 3)][0];
			LED[1][5] = time_number[*(data + 3)][1];
			LED[1][6] = time_number[*(data + 3)][2];
			
			LED[2][0] = time_number[*(data + 4)][0];
			LED[2][1] = time_number[*(data + 4)][1];
			LED[2][2] = time_number[*(data + 4)][2];
			
			LED[2][6] = time_number[*(data + 5)][0];
			LED[2][7] = time_number[*(data + 5)][1];
			LED[3][0] = time_number[*(data + 5)][2];
			
			LED[3][2] = time_number[*(data + 6)][0];
			LED[3][3] = time_number[*(data + 6)][1];
			LED[3][4] = time_number[*(data + 6)][2];
		break;	

		case 0x25:			
		//8*8 horse to up
			for(i = 0 ; i < 4 ; i++)		
			{
				for(j = 0 ; j < 8 ; j++)		
				{
					LED[i][j] = (LED[i][j] >> 1) |  (LED[i][j] << 7) ;
				}						
			}
		break;
			
		case 0x26:			
		//8*8 horse to down
		
			for(i = 0 ; i < 4 ; i++)		
			{
				for(j = 0 ; j < 8 ; j++)		
				{
					LED[i][j] = (LED[i][j] << 1) |  (LED[i][j] >> 7) ;
				}					
			}
		break;
		
		case 0x27:			
		//VB_open
			VB_mod = *(data+1);
		break;
	
		case 0x28:			
		//time HH:MM  or  __:SS
		//H or __(11)
			Matrix_LED(0,0,0x0);
			Matrix_LED(0,1,math_number[*(data+1)][0]);
			Matrix_LED(0,2,math_number[*(data+1)][1]);
			Matrix_LED(0,3,math_number[*(data+1)][2]);
			Matrix_LED(0,4,math_number[*(data+1)][3]);
			Matrix_LED(0,5,math_number[*(data+1)][4]);
			Matrix_LED(0,6,0x0);
			Matrix_LED(0,7,0x0);
			
			//H or __(11)
			Matrix_LED(1,0,math_number[*(data+2)][0]);
			Matrix_LED(1,1,math_number[*(data+2)][1]);
			Matrix_LED(1,2,math_number[*(data+2)][2]);
			Matrix_LED(1,3,math_number[*(data+2)][3]);
			Matrix_LED(1,4,math_number[*(data+2)][4]);
			Matrix_LED(1,5,0x0);
			Matrix_LED(1,6,0x0);	
			
			if (*(data+5) ==1)
			{
			//:: (10)
				Matrix_LED(1,7,math_number[10][0]);
				Matrix_LED(2,0,math_number[10][0]);
			}	
			else
			{
				Matrix_LED(1,7,0x0);
				Matrix_LED(2,0,0x0);
			}
			
			//M  or  S
			Matrix_LED(2,1,0x0);
			Matrix_LED(2,2,0x0);
			Matrix_LED(2,3,math_number[*(data+3)][0]);
			Matrix_LED(2,4,math_number[*(data+3)][1]);
			Matrix_LED(2,5,math_number[*(data+3)][2]);
			Matrix_LED(2,6,math_number[*(data+3)][3]);
			Matrix_LED(2,7,math_number[*(data+3)][4]);
			
			//M  or  S
			Matrix_LED(3,0,0x0);
			Matrix_LED(3,1,0x0);
			Matrix_LED(3,2,math_number[*(data+4)][0]);
			Matrix_LED(3,3,math_number[*(data+4)][1]);
			Matrix_LED(3,4,math_number[*(data+4)][2]);
			Matrix_LED(3,5,math_number[*(data+4)][3]);
			Matrix_LED(3,6,math_number[*(data+4)][4]);
			Matrix_LED(3,7,0x0);
		break;		
	
		default:
		break;		
	} 
	
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0; /* return 1 if this was the last chunk */
	}
}

usbMsgLen_t usbFunctionSetup(uchar data[8])
{
	usbRequest_t    *rq = (void *)data;
    if((rq->bmRequestType & USBRQ_TYPE_MASK) == USBRQ_TYPE_CLASS){    /* HID class request */
        if(rq->bRequest == USBRQ_HID_GET_REPORT){  /* wValue: ReportType (highbyte), ReportID (lowbyte) */
            /* since we have only one report type, we can ignore the report-ID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* use usbFunctionRead() to obtain data */
        }else if(rq->bRequest == USBRQ_HID_SET_REPORT){
            /* since we have only one report type, we can ignore the report-ID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* use usbFunctionWrite() to receive data from host */
        }
    }else{
        /* ignore vendor type requests, we don't use any */
    }
    return 0;
}

int main(void)
{

	uchar i;
    wdt_enable(WDTO_1S);
					// ---------------------------------------------------------
	                //   DDB7   DDB6   DDB5   DDB4   DDB3   DDB2   DDB1   DDB0
	DDRB = 0x3F;    //    0      0      1      1      1      1      1      1
	                //    x      x     out    out    out    out    out    out
					// ---------------------------------------------------------
                    //  PORTB7 PORTB6 PORTB5 PORTB4 PORTB3 PORTB2 PORTB1 PORTB0
	PORTB = 0x00;   //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

					// ---------------------------------------------------------
				    //    -     DDC6   DDC5   DDC4   DDC3   DDC2   DDC1   DDC0
	DDRC = 0x3f;    //    0      0      1      1      1      1      1      1
					//  (Read)  reset   out    out    out    out    out    out
					// ---------------------------------------------------------
                    //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0
    PORTC = 0x40;   //    0      1      0      0      0      0      0      0
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xE2;    //    1      1      1      0      0      0      1      0
					//   out    out    out    D-     in      D+     out    in
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
	PORTD = 0x00;   //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------
					
	Time_Number();
	Math_Number();
	Moving_Logo();
	
    usbInit();
    usbDeviceDisconnect();  // enforce re-enumeration, do this while interrupts are disabled! 
	
	
	i = 0;
		
     while(--i){             // fake USB disconnect for > 250 ms 
        wdt_reset();
        _delay_ms(0.1);
    }	
	
    usbDeviceConnect();
    sei();
	
	/* main event loop */
    while(1)
	{
		wdt_reset();
		usbPoll();	
	
		switch(VB_mod)
		{
		case 0:
			VB_0();
		break;	
		
		case 1:
			VB_1();
		break;	
		
		case 2:
			VB_2();
		break;	
		
		default:
		break;			
		}	
		
		Matrix(0x08,0); 	
		Matrix(0x10,1); 		
		Matrix(0x20,2); 
		Matrix(0x40,3); 
    }
    return 0;
}

/* ------------------------------------------------------------------------- */
