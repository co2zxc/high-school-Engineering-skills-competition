/* Name: main.c
 * Project: hid-data, example how to use HID for data transfer
 * Author: Christian Starkjohann
 * Creation Date: 2008-04-11
 * Tabsize: 4
 * Copyright: (c) 2008 by OBJECTIVE DEVELOPMENT Software GmbH
 * License: GNU GPL v2 (see License.txt), GNU GPL v3 or proprietary (CommercialLicense.txt)
 */

/*
This example should run on most AVRs with only little changes. No special
hardware resources except INT0 are used. You may have to change usbconfig.h for
different I/O pins for USB. Please note that USB D+ must be the INT0 pin, or
at least be connected to INT0 as well.
*/

#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* for sei() */
#include <util/delay.h>     /* for _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* required by usbdrv.h */
#include "usbdrv.h"
#include "oddebug.h"        /* This is also an example for using debug macros */

static uchar j;
static uchar rxHour, rxMinute, rxSecond, rxStNumber;
static uchar rxSecond_Last, flash_Second, flash_CNT1, flash_CNT2;
static uchar left_right;			// 0:Left  1:Right
static uchar step_CNT1, step_CNT2;	// Step time counter1,counter2
static uchar move_CNT;				// Move counter

static uchar T_col, T_hour, T_minute, T_second;
static uchar c,numberCol;
static uchar keyin;

static uchar numbers09[10][5] = 
{
	{0x7E,0x81,0x81,0x81,0x7E},  // 0
	{0x00,0x82,0xFF,0x80,0x00},  // 1
	{0xC2,0xA1,0x91,0x89,0x86},  // 2
	{0x41,0x81,0x89,0x95,0x63},  // 3
	{0x38,0x24,0x22,0xFF,0x20},  // 4
	{0x4F,0x89,0x89,0x89,0x71},  // 5
	{0x7C,0x92,0x91,0x91,0x60},  // 6
	{0x01,0x01,0xF9,0x05,0x03},  // 7
	{0x66,0x99,0x91,0x99,0x66},  // 8
	{0x06,0x89,0x89,0x49,0x3E}   // 9
};

static uchar column[32] =
{
	0x08,0x09,0x0A,0x0B,0x0C,0x0D,0x0E,0x0F,  // column  0 to  7
	0x10,0x11,0x12,0x13,0x14,0x15,0x16,0x17,  // column  8 to 15
	0x20,0x21,0x22,0x23,0x24,0x25,0x26,0x27,  // column 16 to 23
	0x40,0x41,0x42,0x43,0x44,0x45,0x46,0x47   // column 24 to 31
};

static uchar logo[64] =
{
	0x3F,0xA1,0xE1,0xE1,0xE1,0xE1,0xA1,0x3F,  		// column  0 to  7
	0x00,0x06,0x05,0x09,0x91,0xE1,0xE1,0x91,  		// column  8 to 15
	0x09,0x05,0x06,0x00,0xC0,0xC0,0xF8,0x88,  		// column 16 to 23
	0x8F,0x81,0x8F,0x88,0xF8,0xC0,0xC0,0x00,   		// column 24 to 31
	0x3F,0xA1,0xE1,0xE1,0xE1,0xE1,0xA1,0x3F,  		// column  0 to  7
	0x00,0x06,0x05,0x09,0x91,0xE1,0xE1,0x91,  		// column  8 to 15
	0x09,0x05,0x06,0x00,0xC0,0xC0,0xF8,0x88,  		// column 16 to 23
	0x8F,0x81,0x8F,0x88,0xF8,0xC0,0xC0,0x00   		// column 24 to 31
};

static void outputByte(uchar b)
{
    PORTB = (PORTB & ~0x3f) | (b & 0x3f);
    PORTD = (PORTD & ~0xc0) | (b & 0xc0);
}

static void outputControl(uchar b)
{
    PORTD = (PORTD & 0xdf) | (b & 0x20);
}

static void outputDotMatrixCol(uchar b)
{
    PORTC = (PORTC & ~0x7f) | (b & 0x7f);
}

static uchar inputKeys(void)
{
	uchar	b;
	b = PIND & 0x03;
    return b;
}

static void showCurrentTime(uchar col, uchar hour, uchar minute, uchar second)
{
	T_col = col;
	T_hour = hour;
	T_minute = minute;
	T_second = second;
	
	if ((T_col >= 1) && (T_col <= 5)) {
		if (T_hour <= 23) {
			if (T_hour >= 10)
				c = T_hour / 10;
			else
				c = 0;
			numberCol = T_col - 1;
			outputByte(0);
			outputDotMatrixCol(column[T_col]);
			outputByte(numbers09[c][numberCol]);
		} else {
			outputByte(0);
			outputDotMatrixCol(0);
			return;
		}
	} else if ((T_col == 0)||(T_col == 6)||(T_col == 7)||(T_col == 13)||(T_col == 14)||(T_col == 17)
				||(T_col == 18)||(T_col == 24)||(T_col == 25)||(T_col == 31)) {
		outputByte(0);
		outputDotMatrixCol(0);
		return;
	} else if ((T_col >= 8) && (T_col <= 12)) {
		if (T_hour <= 23) {
			c = T_hour % 10;
			numberCol = T_col - 8;
			outputByte(0);
			outputDotMatrixCol(column[T_col]);
			outputByte(numbers09[c][numberCol]);
		} else {
			outputByte(0);
			outputDotMatrixCol(0);
			return;
		}
	} else if ((T_col == 15)||(T_col == 16)) {
		outputByte(0);
		outputDotMatrixCol(column[T_col]);
		if (rxSecond != rxSecond_Last) {
			flash_Second = 1;					// Flash Second ON
			rxSecond_Last = rxSecond;
		} else if ((flash_Second == 3) && (rxSecond == rxSecond_Last)) {
			outputByte(0x66);
		}
		if (flash_Second == 1) {
			flash_CNT1 = flash_CNT1 + 1;
			if (flash_CNT1 >= 200) {
				flash_CNT1 = 0;
				flash_CNT2 = flash_CNT2 + 1;
				if (flash_CNT2 >= 8) {
					flash_CNT2 = 0;
					flash_Second = 2;
				}
			}
			outputByte(0x66);
		} else if (flash_Second == 2) {
			flash_CNT1 = flash_CNT1 + 1;
			if (flash_CNT1 >= 200) {
				flash_CNT1 = 0;
				flash_CNT2 = flash_CNT2 + 1;
				if (flash_CNT2 >= 6) {
					flash_CNT2 = 0;
					flash_Second = 3;
				}
			}
		}
		return;
	} else if ((T_col >= 19) && (T_col <= 23)) {
		if (T_minute <= 59) {
			if (T_minute >= 10)
				c = T_minute / 10;
			else
				c = 0;
			numberCol = T_col - 19;
			outputByte(0);
			outputDotMatrixCol(column[T_col]);
			outputByte(numbers09[c][numberCol]);
		} else {
			outputByte(0);
			outputDotMatrixCol(0);
			return;
		}
	} else if ((T_col >= 26) && (T_col <= 30)) {
		if (T_minute <= 59) {
			c = T_minute % 10;
			numberCol = T_col - 26;
			outputByte(0);
			outputDotMatrixCol(column[T_col]);
			outputByte(numbers09[c][numberCol]);
		} else {
			outputByte(0);
			outputDotMatrixCol(0);
			return;
		}
	}
}

static void showCurrentSecond(uchar col, uchar second)
{
	uchar c,numberCol;
	if (((col >= 0) && (col <= 14))||((col >= 17) && (col <= 18))
	      ||((col >= 24) && (col <= 25))||(col == 31)) {
		outputByte(0);
		outputDotMatrixCol(0);
		return;
	} else if ((col >= 15) && (col <= 16)) {
		outputByte(0);
		outputDotMatrixCol(column[col]);
		outputByte(0x66);
		return;
	} else if ((col >= 19) && (col <= 23)) {
		if (second <= 59) {
			if (second >= 10)
				c = second / 10;
			else
				c = 0;
			numberCol = col - 19;
			outputByte(0);
			outputDotMatrixCol(column[col]);
			outputByte(numbers09[c][numberCol]);
		} else {
			outputByte(0);
			outputDotMatrixCol(0);
			return;
		}
	} else if ((col >= 26) && (col <= 30)) {
		if (second <= 59) {
			c = second % 10;
			numberCol = col - 26;
			outputByte(0);
			outputDotMatrixCol(column[col]);
			outputByte(numbers09[c][numberCol]);
		} else {
			outputByte(0);
			outputDotMatrixCol(0);
			return;
		}
	}
}

static void showStationNumber(uchar col, uchar stNumber)
{
	uchar c,numberCol;
	if (((col >= 0) && (col <= 18))||((col >= 24) && (col <= 25))||(col == 31)) {
		outputByte(0);
		outputDotMatrixCol(0);
		return;
	} else if ((col >= 19) && (col <= 23)) {
		if (stNumber <= 99) {
			if (stNumber >= 10) {
				c = stNumber / 10;
			} else {
				c = 0;
			}
			numberCol = col - 19;
			outputByte(0);
			outputDotMatrixCol(column[col]);
			outputByte(numbers09[c][numberCol]);
		} else {
			outputByte(0);
			outputDotMatrixCol(0);
			return;
		}
	} else if ((col >= 26) && (col <= 30)) {	
		if (stNumber <= 99) {
			if (stNumber >= 10) {
				c = stNumber % 10;
			} else {
				c = stNumber;
			}
			numberCol = col - 26;
			outputByte(0);
			outputDotMatrixCol(column[col]);
			outputByte(numbers09[c][numberCol]);
		} else {
			outputByte(0);
			outputDotMatrixCol(0);
			return;
		}
	}
}
	
/* ------------------------------------------------------------------------- */
/* ----------------------------- USB interface ----------------------------- */
/* ------------------------------------------------------------------------- */

PROGMEM const char usbHidReportDescriptor[22] = {    /* USB report descriptor */
    0x06, 0x00, 0xff,              // USAGE_PAGE (Generic Desktop)
    0x09, 0x01,                    // USAGE (Vendor Usage 1)
    0xa1, 0x01,                    // COLLECTION (Application)
    0x15, 0x00,                    //   LOGICAL_MINIMUM (0)
    0x26, 0xff, 0x00,              //   LOGICAL_MAXIMUM (255)
    0x75, 0x08,                    //   REPORT_SIZE (8)
    0x95, 0x80,                    //   REPORT_COUNT (128)
    0x09, 0x00,                    //   USAGE (Undefined)
    0xb2, 0x02, 0x01,              //   FEATURE (Data,Var,Abs,Buf)
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
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		*(data) = keyin;	// Key value
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
	static uchar tempB,tempD;
	
    if(bytesRemaining == 0)
        return 1;               /* end of transfer */
    if(len > bytesRemaining)
        len = bytesRemaining;
    //eeprom_write_block(data, (uchar *)0 + currentAddress, len);
	if(currentAddress == 0){

		tempB = PORTB;			// Save the original data of PORTB
        tempD = PORTD;			// Save the original data of PORTD
		outputDotMatrixCol(0);	// Blank Dot Matrix
		outputByte(0);        	// Blank Dot Matrix

		outputByte(*(data));		// Output display data
		
		if((*(data+1) & 0x20) == 0x20) {
			outputControl(0x20);		// ---- Generate a Positive pulse
			outputControl(0x0);			// ----
		} else {
			outputControl(0x0);
		}

		rxHour= *(data+2);			// Hour from Host
		rxMinute = *(data+3);		// Minute from Host
		rxSecond = *(data+4);		// Second from Host
		rxStNumber = *(data+5);		// Station Number from Host

		PORTB = tempB;			// Restore the original data of PORTB
		PORTD = tempD;			// Restore the original data of PORTD
		outputDotMatrixCol(j);
		
	}
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0; /* return 1 if this was the last chunk */
}

/* ------------------------------------------------------------------------- */

usbMsgLen_t usbFunctionSetup(uchar data[8])
{
usbRequest_t    *rq = (void *)data;

    if((rq->bmRequestType & USBRQ_TYPE_MASK) == USBRQ_TYPE_CLASS){    /* HID class request */
        if(rq->bRequest == USBRQ_HID_GET_REPORT){  /* wValue: ReportType (highbyte), ReportID (lowbyte) */
            /* since we have only one report type, we can ignore the report-ID */
            bytesRemaining = 8;
			//bytesRemaining = 128;
            currentAddress = 0;
            return USB_NO_MSG;  /* use usbFunctionRead() to obtain data */
        }else if(rq->bRequest == USBRQ_HID_SET_REPORT){
            /* since we have only one report type, we can ignore the report-ID */
            bytesRemaining = 8;
			//bytesRemaining = 128;
            currentAddress = 0;
            return USB_NO_MSG;  /* use usbFunctionWrite() to receive data from host */
        }
    }else{
        /* ignore vendor type requests, we don't use any */
    }
    return 0;
}

/* ------------------------------------------------------------------------- */

int main(void)
{	
	uchar   i;
    wdt_enable(WDTO_1S);
    /* Even if you don't use the watchdog, turn it off here. On newer devices,
     * the status of the watchdog (on/off, period) is PRESERVED OVER RESET!
     */

					// ---------------------------------------------------------
	                //   DDB7   DDB6   DDB5   DDB4   DDB3   DDB2   DDB1   DDB0
	DDRB = 0x3f;    //    0      0      1      1      1      1      1      1
	                //    x      x     out    out    out    out    out    out
					// ---------------------------------------------------------
                    //  PORTB7 PORTB6 PORTB5 PORTB4 PORTB3 PORTB2 PORTB1 PORTB0
	PORTB = 0x0;    //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

					// ---------------------------------------------------------
				    //    -     DDC6   DDC5   DDC4   DDC3   DDC2   DDC1   DDC0
	DDRC = 0x7f;    //    0      1      1      1      1      1      1      1  
					//  (Read)  out    out    out    out    out    out    out							 
					// ---------------------------------------------------------
                    //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0							 
    PORTC = 0x0;    //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xe0;    //    1      1      1      0      0      0      0      0    
					//   out    out    out    D-     in     D+     in     in
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
	PORTD = 0x0;    //    0      0      0      0      0      0      0      0  
					// ---------------------------------------------------------
	
    usbInit();
    usbDeviceDisconnect();  /* enforce re-enumeration, do this while interrupts are disabled! */
	i = 0;
    while(--i){             /* fake USB disconnect for > 250 ms */
        wdt_reset();
        _delay_ms(1);
    }
    usbDeviceConnect();
    sei();
	
	j = 0;
	rxHour = 23;
	rxMinute = 59;
	rxSecond = 47;
	rxStNumber = 98;
	flash_Second = 0;		// Flash Second OFF
	left_right = 0;			// Move Left
	step_CNT1 = 0;
	step_CNT2 = 0;
	move_CNT = 0;

    for(;;){                /* main event loop */
        wdt_reset();
        usbPoll();
		
		step_CNT1 = step_CNT1 + 1;
		if (step_CNT1 >= 200) {
			step_CNT1 = 0;
			step_CNT2 = step_CNT2 + 1;
		}
		if (step_CNT2 >= 100) {
			step_CNT2 = 0;
			if (left_right == 0) {			// Move Left
				move_CNT = move_CNT + 1;
				if (move_CNT == 32) {
					if (left_right == 0) {
						left_right = 1;
					} else {
						left_right = 0;
					}
				}
			} else {						// Move Right
				move_CNT = move_CNT - 1;
				if (move_CNT == 0) {
					if (left_right == 1) {
						left_right = 0;
					} else {
						left_right = 1;
					}
				}
			}
		}

		keyin = inputKeys();

		outputDotMatrixCol(0);
		outputByte(0);
		
		switch(keyin)
		{
			case 0x00:
				outputByte(0);
				outputDotMatrixCol(column[j]);
				outputByte(logo[j + move_CNT]);
				break;
				
			case 0x01:
				showStationNumber(j, rxStNumber);
				break;
				
			case 0x02:
				showCurrentSecond(j, rxSecond);
				break;
				
			default:
				showCurrentTime(j, rxHour, rxMinute, rxSecond);
				break;
		}
		if (j==31)  j=0;  else  j++;
    }
    return 0;
}

/* ------------------------------------------------------------------------- */
