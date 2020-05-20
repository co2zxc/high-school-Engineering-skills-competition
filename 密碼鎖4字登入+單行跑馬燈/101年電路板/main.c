#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* for sei() */
#include <util/delay.h>     /* for _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* required by usbdrv.h */
#include "usbdrv.h"
#include "oddebug.h"        /* This is also an example for using debug macros */

static uchar keyin;
static uchar LED[4][8];
static uchar timenumber[11][3];
static uchar number[10][8];
static	double nb ;
static	double nb2 ;

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

static uchar inputKeys()
{
	uchar b;
	b = PIND & 0x01;
	return b;
}

static void Matrix(uchar Row, uchar Data)
{
	//8*8 LED
	uchar i;
		wdt_reset();
		usbPoll();
		for(i = 0 ; i < 8 ; i++)
		{
			PORTB =(PORTB & 0xC0) | (LED[Data][i] & 0x3f); //0xC0 11000000 //0x3f 00111111
			PORTD =(PORTD & 0x3F) | (LED[Data][i] & 0xC0); //0x3F 00111111 //0xC0 11000000  
			outputEnable(Row + i);
			_delay_ms(0.01);
			outputEnable(0x00);
		}	
}

static void LEDclear()
{
			outputLED(0x0); // outputLED(*(data+1))
			outputControl();	
			LED[0][0] = 0x0;
			LED[0][1] = 0x0;	
			LED[0][2] = 0x0;
			LED[0][3] = 0x0;
			LED[0][4] = 0x0;
			LED[0][5] = 0x0;
			LED[0][6] = 0x0;
			LED[0][7] = 0x0;

			LED[1][0] = 0x0;
			LED[1][1] = 0x0;	
			LED[1][2] = 0x0;
			LED[1][3] = 0x0;
			LED[1][4] = 0x0;
			LED[1][5] = 0x0;
			LED[1][6] = 0x0;
			LED[1][7] = 0x0;

			LED[2][0] = 0x0;
			LED[2][1] = 0x0;	
			LED[2][2] = 0x0;
			LED[2][3] = 0x0;
			LED[2][4] = 0x0;
			LED[2][5] = 0x0;
			LED[2][6] = 0x0;
			LED[2][7] = 0x0;
		
			LED[3][0] = 0x0;
			LED[3][1] = 0x0;	
			LED[3][2] = 0x0;
			LED[3][3] = 0x0;
			LED[3][4] = 0x0;
			LED[3][5] = 0x0;
			LED[3][6] = 0x0;
			LED[3][7] = 0x0;
			
}

static void LEDset()
{
				LED[0][0] = 0xFF;
				LED[0][1] = 0xFF;
				LED[0][2] = 0xFF;
				LED[0][3] = 0xFF;
				LED[0][4] = 0xFF;
				LED[0][5] = 0xFF;
				LED[0][6] = 0xFF;
				LED[0][7] = 0xFF;
				
				LED[1][0] = 0xFF;
				LED[1][1] = 0xFF;
				LED[1][2] = 0xFF;
				LED[1][3] = 0xFF;
				LED[1][4] = 0xFF;
				LED[1][5] = 0xFF;
				LED[1][6] = 0xFF;
				LED[1][7] = 0xFF;
				
				LED[2][0] = 0xFF;
				LED[2][1] = 0xFF;
				LED[2][2] = 0xFF;
				LED[2][3] = 0xFF;
				LED[2][4] = 0xFF;
				LED[2][5] = 0xFF;
				LED[2][6] = 0xFF;
				LED[2][7] = 0xFF;

				
				LED[3][0] = 0xFF;
				LED[3][1] = 0xFF;
				LED[3][2] = 0xFF;
				LED[3][3] = 0xFF;
				LED[3][4] = 0xFF;
				LED[3][5] = 0xFF;
				LED[3][6] = 0xFF;
				LED[3][7] = 0xFF;
				
		
				outputLED(0xFF); // outputLED(*(data+1))
				outputControl();
}

static void LEDset2()
{
				LED[0][0] = 0x8;
				LED[0][1] = 0x8;
				LED[0][2] = 0x8;
				LED[0][3] = 0x8;
				LED[0][4] = 0x8;
				LED[0][5] = 0x8;
				LED[0][6] = 0x8;
				LED[0][7] = 0x8;
				
				LED[1][2] = number[9][2];
				LED[1][3] = number[9][3];
				LED[1][4] = number[9][4];
				LED[1][5] = number[9][5];
				
				LED[2][2] = number[9][2];
				LED[2][3] = number[9][3];
				LED[2][4] = number[9][4];
				LED[2][5] = number[9][5];
				
				LED[3][0] = 0x8;
				LED[3][1] = 0x8;
				LED[3][2] = 0x8;
				LED[3][3] = 0x8;
				LED[3][4] = 0x8;
				LED[3][5] = 0x8;
				LED[3][6] = 0x8;
				LED[3][7] = 0x8;
			
				outputLED(25); // outputLED(*(data+1))
				outputControl();
}

static void numberset()
{
//timenumber 0
timenumber[0][0] = 0x7F;
timenumber[0][1] = 0x41;
timenumber[0][2] = 0x7F;
//timenumber 1
timenumber[1][0] = 0x42;
timenumber[1][1] = 0x7F;
timenumber[1][2] = 0x40;
//timenumber 2
timenumber[2][0] = 0x79;
timenumber[2][1] = 0x49;
timenumber[2][2] = 0x4F;
//timenumber 3
timenumber[3][0] = 0x49;
timenumber[3][1] = 0x49;
timenumber[3][2] = 0x7F;
//timenumber 4
timenumber[4][0] = 0xF;
timenumber[4][1] = 0x8;
timenumber[4][2] = 0x7F;
//timenumber 5
timenumber[5][0] = 0x4F;
timenumber[5][1] = 0x49;
timenumber[5][2] = 0x79;
//timenumber 6
timenumber[6][0] = 0x7F;
timenumber[6][1] = 0x49;
timenumber[6][2] = 0x79;
//timenumber 7
timenumber[7][0] = 0x7;
timenumber[7][1] = 0x1;
timenumber[7][2] = 0x7F;
//timenumber 8
timenumber[8][0] = 0x7F;
timenumber[8][1] = 0x49;
timenumber[8][2] = 0x7F;
//timenumber 9
timenumber[9][0] = 0xF;
timenumber[9][1] = 0x9;
timenumber[9][2] = 0x7F;
//:
timenumber[10][0] = 0x66;


number[0][2] = 0x1F;
number[0][3] = 0x11;
number[0][4] = 0x11;
number[0][5] = 0x1F;

number[1][5] = 0x1F;

number[2][2] = 0x1D;
number[2][3] = 0x15;
number[2][4] = 0x15;
number[2][5] = 0x17;

number[3][2] = 0x15;
number[3][3] = 0x15;
number[3][4] = 0x15;
number[3][5] = 0x1F;

number[4][2] = 0x7;
number[4][3] = 0x4;
number[4][4] = 0x4;
number[4][5] = 0x1F;

number[5][2] = 0x17;
number[5][3] = 0x15;
number[5][4] = 0x15;
number[5][5] = 0x1D;

number[6][2] = 0x1F;
number[6][3] = 0x14;
number[6][4] = 0x14;
number[6][5] = 0x1C;

number[7][2] = 0x7;
number[7][3] = 0x1;
number[7][4] = 0x1;
number[7][5] = 0x1F;

number[8][2] = 0x1F;
number[8][3] = 0x15;
number[8][4] = 0x15;
number[8][5] = 0x1F;

number[9][2] = 0x17;
number[9][3] = 0x15;
number[9][4] = 0x15;
number[9][5] = 0x1F;
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
		*(data) = inputKeys();	// Key value
		*(data+1) = 0;		// Reserve (zero)
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

		//8*8 LED
		if(*(data) == 0x20) {
			LED[*(data + 1)][*(data + 2)] = *(data + 3);
		}
		//8 bit LED
		else if(*(data) == 0x21)
		{
			outputLED(data[1]); // outputLED(*(data+1))
			outputControl();			// Clock
		}
		//8*8 horse to left
		else if(*(data) == 0x22)
		{
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
			LED[3][7] = *(data + 1);
		}
		//8*8 horse to right
		else if(*(data) == 0x23)
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
			LED[0][0] = *(data + 1);
		}
		else if(*(data) == 0x24)
		{
			LED[0][2] = timenumber[*(data + 1)][0];
			LED[0][3] = timenumber[*(data + 1)][1];
			LED[0][4] = timenumber[*(data + 1)][2];
			
			LED[0][6] = timenumber[*(data + 2)][0];
			LED[0][7] = timenumber[*(data + 2)][1];
			LED[1][0] = timenumber[*(data + 2)][2];
			
			LED[1][4] = timenumber[*(data + 3)][0];
			LED[1][5] = timenumber[*(data + 3)][1];
			LED[1][6] = timenumber[*(data + 3)][2];
			
			LED[2][0] = timenumber[*(data + 4)][0];
			LED[2][1] = timenumber[*(data + 4)][1];
			LED[2][2] = timenumber[*(data + 4)][2];
			
			LED[2][6] = timenumber[*(data + 5)][0];
			LED[2][7] = timenumber[*(data + 5)][1];
			LED[3][0] = timenumber[*(data + 5)][2];
			
			LED[3][2] = timenumber[*(data + 6)][0];
			LED[3][3] = timenumber[*(data + 6)][1];
			LED[3][4] = timenumber[*(data + 6)][2];
		}
		else if(*(data) == 0x25)
		{
			PORTD = (PORTD & 0xF7)  |  *(data + 1)  ; //0xF7 11110111
		}

	}
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0; /* return 1 if this was the last chunk */
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
	numberset();

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
	DDRD = 0xEA;    //    1      1      1      0      1      0      1      0
					//   out    out    out    D-     out    D+     out    in
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
	PORTD = 0x00;   //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

    usbInit();
    usbDeviceDisconnect();  /* enforce re-enumeration, do this while interrupts are disabled! */
	i = 0;
	
	nb = 0;
	nb2 = 0;
    while(--i){             /* fake USB disconnect for > 250 ms */
        wdt_reset();
       _delay_ms(0.1);		
    }
    usbDeviceConnect();
    sei();
		
    while(1)
	{   
		if (nb2 < 3){
		
			if (nb ==0){
			LEDset();	
			}
			if (nb ==1000){
			LEDclear();	
			}
			if (nb ==2000){
				nb = 0;	
				nb2 = nb2 +1 ;						
			} else {
				nb = nb +1 ;
			}
			if (nb2 ==3){
				LEDset2();			
			}
		}
		
		
		
		wdt_reset();
		usbPoll();		
		Matrix(0x08,0); 	
		Matrix(0x10,1); 		
		Matrix(0x20,2); 
		Matrix(0x40,3); 
		
	}
	   return 0;
}	
    
 


/* ------------------------------------------------------------------------- */
