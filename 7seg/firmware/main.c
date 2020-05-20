#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* for sei() */
#include <util/delay.h>     /* for _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* required by usbdrv.h */
#include "usbdrv.h"
#include "oddebug.h"        /* This is also an example for using debug macros */

static uchar d0,d1,d2,d3,porta,portb,portc;
static uchar seg[16]={0xc0,0xf9,0xa4,0xb0,0x99,0x92,0x82,0xf8,0x80,0x98,0x83,0xc6,0xa1,0x86,0x8e,0xbf};
static void inputPORTA(void)
{
	uchar b,d;
	b = PINB;
	d = PIND;
	porta = (b & 0x3f) | (d & 0xc0);
}

static void inputPORTB(void)
{
	uchar c,d;
	c = PINC;
	d = PIND;
	portb = (c & 0x3f) | ((d << 6) & 0xc0);
}

static void inputPORTC(void)
{
	uchar b,c,d;
	c = PINC;
	d = PIND;
	b = ((d >> 2) & 0x02) | ((d >> 3) & 0x04);
	portc = (b | ((c >> 6) & 0x01));
}

//static uchar inputPORTX()
static void inputPORTX(void)
{
	uchar b,c,d;
	b = PINB;
	c = PINC;
	d = PIND;
	porta = (b & 0x3f) | (d & 0xc0);
	portb = (c & 0x3f) | ((d << 6) & 0xc0);
	portc = ((d >> 2) & 0x02) | ((d >> 3) & 0x04);
	portc = (portc | ((c >> 6) & 0x01));
//	return 0;
}

static void outputPORTA(uchar a)
{
	PORTB = (PORTB & ~0x3f) | (a & 0x3f);
	PORTD = (PORTD & ~0xc0) | (a & 0xc0);
}

static void outputPORTB(uchar b)	
{
    PORTC = (PORTC & ~0x3f) | (b & 0x3f);
    PORTD = (PORTD & ~0x03) | ((b >> 6) & 0x03);
}

void outputPORTC(uchar c)	
{
    PORTC = (PORTC & ~0x40) | ((c << 6) & 0x40);
    PORTD = (PORTD & ~0x08) | ((c << 2) & 0x08);
	PORTD = (PORTD & ~0x20) | ((c << 3) & 0x20);
}

void outputPORTX(uchar a,uchar b,uchar c)	
{
	PORTB = (PORTB & ~0x3f) | (a & 0x3f);
	PORTD = (PORTD & ~0xc0) | (a & 0xc0);
    PORTC = (PORTC & ~0x3f) | (b & 0x3f);
    PORTD = (PORTD & ~0x03) | ((b >> 6) & 0x03);
    PORTC = (PORTC & ~0x40) | ((b << 6) & 0x40);
    PORTD = (PORTD & ~0x08) | ((b << 2) & 0x08);
	PORTD = (PORTD & ~0x20) | ((b << 3) & 0x20);
}

static void outputPulse(void)
{
	PORTC = (PORTC & 0xbf);
	PORTC = (PORTC | 0x40);
	PORTC = (PORTC & 0xbf);
}

static void ctrl(uchar p)
{
    p = p & 0x8f;
	switch (p)
    {
        case 0x81:	outputPulse();
					break;
        case 0x88:	outputPORTC(0x00);
					break;
        case 0x89:	outputPORTC(0x00);
					break;
        case 0x8a:	outputPORTC(0x02);
					break;
        case 0x8b:	outputPORTC(0x03);
					break;
        case 0x8c:	outputPORTC(0x04);
					break;
        case 0x8d:	outputPORTC(0x05);
					break;
        case 0x8e:	outputPORTC(0x06);
					break;
        case 0x8f:	outputPORTC(0x07);
		default:	break;
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

/* usbFunctionRead() is called when the host requests a chunk of data from
 * the device. For more information see the documentation in usbdrv/usbdrv.h.
 */
uchar   usbFunctionRead(uchar *data, uchar len)
{
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		*(data) = porta;	// PORTA input
		*(data+1) = portb;	// PORTB input
		*(data+2) = portc;	// PORTC input
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
    uchar a,b,c,p,tmp1,tmp2,tmp3;
	
    if(bytesRemaining == 0)
        return 1;               /* end of transfer */
    if(len > bytesRemaining)
        len = bytesRemaining;
		
	if(currentAddress == 0){
		if(*(data) == 0x01) {
			inputPORTA();							// input PORTA
		}
		if(*(data) == 0x02) {
			inputPORTB();							// input PORTB
		}
		if(*(data) == 0x03) {
			inputPORTC();							// input PORTC
		}
		if(*(data) == 0x04) {
			inputPORTX();							// input PORTX
		}
		if(*(data) == 0x11) {
			a=*(data + 1);		p=*(data + 7);		// out PORTA
			outputPORTA(a);
			ctrl(p);
		}
		if(*(data) == 0x12) {
			b=*(data + 1);		p=*(data + 7);		// out PORTB
			outputPORTB(b);
			ctrl(p);
		}
		if(*(data) == 0x13) {
			c = *(data + 1);		p=*(data + 7);		// out PORTC
			outputPORTC(c);
			ctrl(p);
		}
		if(*(data) == 0x14) {
			a=*(data + 1); b=*(data + 2); c=*(data + 3); p=*(data + 7);		// out PORTX
			outputPORTX(a,b,c);
			ctrl(p);
		}
		if(*(data) == 0x20) {	// PORTX every Bit => input:0  output:1
			a = *(data + 1);
			b = *(data + 2);
			c = *(data + 3);
            DDRB = (DDRB & 0xc0) | (a & 0x3f);
			tmp1 = ((c << 6) & 0x40) | (b & 0x3f);
			DDRC = ((DDRC & 0x80) | tmp1);
			tmp2 = ((b >> 6) & 0x03) | ((c << 2) & 0x08);
	        tmp3 = ((c << 3) & 0x20) | (a & 0xc0);
	        DDRD = tmp3 | tmp2;
		}
		if(*(data) == 0x21) {
			d0 = *(data + 1);
			d1 = *(data + 2);
			d2 = *(data + 3);
			d3 = *(data + 4);
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
	uchar i;
    wdt_enable(WDTO_1S);
					// ---------------------------------------------------------
	                //   DDB7   DDB6   DDB5   DDB4   DDB3   DDB2   DDB1   DDB0
	DDRB = 0x3f;    //    0      0      1      1      1      1      1      1
	                //    x      x     out    out    out    out    out    out
					// ---------------------------------------------------------
                    //  PORTB7 PORTB6 PORTB5 PORTB4 PORTB3 PORTB2 PORTB1 PORTB0
//	PORTB = 0x00;   //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

					// ---------------------------------------------------------
				    //    -     DDC6   DDC5   DDC4   DDC3   DDC2   DDC1   DDC0
	DDRC = 0x7f;    //    0      1      1      1      1      1      1      1  
					//  (Read)  out    out    out    out    out    out    out							 
					// ---------------------------------------------------------
                    //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0							 
//  PORTC = 0x40;   //    0      1      0      0      0      0      0      0
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xeb;    //    1      1      1      0      1      0      1      1    
					//   out    out    out    D-     out    D+     out    in
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
//	PORTD = 0x00;   //    0      0      0      0      0      0      0      0  
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

    while(1)
	{                /* main event loop */
	    wdt_reset();
		usbPoll();
		_delay_ms(1);
		
		outputPORTB(0xfe);
		outputPORTA(~seg[d0]);
		_delay_ms(4);
		
		outputPORTB(0xfd);
		outputPORTA(~seg[d1]);
		_delay_ms(4);	
		
		outputPORTB(0xfb);
		outputPORTA(~seg[d2]);
		_delay_ms(4);	
		
		outputPORTB(0xf7);
		outputPORTA(~seg[d3]);
		_delay_ms(4);	
				
    }
    return 0;
}
