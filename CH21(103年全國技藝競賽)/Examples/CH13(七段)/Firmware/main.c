/*   ���عϮ�: USB�����]�p�P���γ]�p�J������d�ҵ{��-CH13(�\�éM�s��) */
/*   �{���\��: �Q��PB��PD�X��4���C�q��ܾ�                             */   

#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* sei():�P�त�_ */
#include <util/delay.h>     /*  _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* usbdrv.h�һݪ��ѼƩw�q */
#include "usbdrv.h"
#include "oddebug.h"        /* �ϥΰ����������d�� */
/* �C�q��ܾ� 0~F�ƭȪ��ѽX�ƭ� */ 
/* PB�ѽX�ƭ� */
unsigned char Led_B[16] = {0x3F,0x06,0x1B,0x0F,0x26,0x2D,0x3D,0x27,0x3F,0x27,0x37,0x3C,0x39,0x1E,0x39,0x31};
/* PD�ѽX�ƭ� */
unsigned char Led_D[16] = {0x00,0x00,0x40,0x40,0x40,0x40,0x40,0x00,0x40,0x40,0x40,0x40,0x00,0x40,0x40,0x40};
unsigned char SEG_Data[4] = {0,0,0,0};
/* �P���PB�PPD��X����C�q��ܾ�����ܪ���� */
static void SEG_Control(uchar b)
{
	PORTD = (PORTD & ~0xC0) | (b & 0xC0);
	PORTB = (PORTB & ~0x3F) | (b & 0x3F);
}
/* �C�q��ܾ��禡�A�|���C�q��ܾ�(U4�AU5�AU6�PU7)�ӧO���� */
static void SEG(uchar w, uchar x, uchar y, uchar z)
{
	PORTD = ((PORTD & ~0x2B) | 0x01);		/* U4 */
	SEG_Control(Led_B[w]|Led_D[w]);
	_delay_ms(1);
	PORTD = ((PORTD & ~0x2B) | 0x02);		/* U5 */
	SEG_Control(Led_B[x]|Led_D[x]);
	_delay_ms(1);
	PORTD = ((PORTD & ~0x2B) | 0x08);		/* U6 */
	SEG_Control(Led_B[y]|Led_D[y]);
	_delay_ms(1);
	PORTD = ((PORTD & ~0x2B) | 0x20);		/* U7 */
	SEG_Control(Led_B[z]|Led_D[z]);
	_delay_ms(1);
}

/* ------------------------------------------------------------------------- */
/* ------------------------------ USB ���� --------------------------------- */
/* ------------------------------------------------------------------------- */

PROGMEM const char usbHidReportDescriptor[22] = {   /* USB ���i�y�z�� */
    0x06, 0x00, 0xff,              // USAGE_PAGE (�c��Ӧ۩w)
    0x09, 0x01,                    // USAGE (Vendor Usage 1)
    0xa1, 0x01,                    // COLLECTION (Application)
    0x15, 0x00,                    // LOGICAL_MINIMUM (0)
    0x26, 0xff, 0x00,              // LOGICAL_MAXIMUM (255)
    0x75, 0x08,                    // REPORT_SIZE (8)
    0x95, 0x08,                    // REPORT_COUNT (8)
    0x09, 0x00,                    // USAGE (Undefined)
    0xb2, 0x02, 0x01,              // FEATURE (Data,Var,Abs,Buf)
    0xc0                           // END_COLLECTION
};

/* �H�U���ܼƬO�ΨӼȦs�ثe��ƶǿ骺���A */
static uchar    currentAddress;
static uchar    bytesRemaining;

/* usbFunctionRead():��D���n�qUSB�˸m�n�D�j���ƮɴN�|�I�susbFunctionRead()*/ 
/* �����T���i�H�i�H�Ѧ�usbdrv/usbdrv.h���*/
uchar   usbFunctionRead(uchar *data, uchar len)
{
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
	}
    currentAddress += len;
    bytesRemaining -= len;
    return len;
}

/* usbFunctionWrite():��D���n�e�X�j���Ƶ�USB�˸m�ɴN�|�I�susbFunctionWrite()*/
/* �����T���i�H�i�H�Ѧ�usbdrv/usbdrv.h���*/
 
uchar   usbFunctionWrite(uchar *data, uchar len)
{
    if(bytesRemaining == 0)
        return 1;               /* USB�ǿ鵲�� */
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0)
	{	
		/* ��X�P����SEG����Ʀ줸�� */
		SEG_Data[0] = data[0];	
		SEG_Data[1] = data[1];
		SEG_Data[2] = data[2];
		SEG_Data[3] = data[3];
	}
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0; /* �p�G�o�O�̫�@���ơA�N�^��1 */
}

/* USB HID�s�խn�D: GET_REPORT�PSET_REPORT */
usbMsgLen_t usbFunctionSetup(uchar data[8])
{
	usbRequest_t    *rq = (void *)data;
    if((rq->bmRequestType & USBRQ_TYPE_MASK) == USBRQ_TYPE_CLASS){    /* HID class request */
        if(rq->bRequest == USBRQ_HID_GET_REPORT){  /* wValue: ReportType (highbyte), ReportID (lowbyte) */
            /* �ѩ�Ȧ��@�ӳ��i�����A�]���������iID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* �ϥ�usbFunctionRead()�Ӷǰe��ƨ�PC�D��*/
        }else if(rq->bRequest == USBRQ_HID_SET_REPORT){
            /* �ѩ�Ȧ��@�ӳ��i�����A�]���������iID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* �ϥ� usbFunctionWrite()�qPC�D���������*/
        }
    }else{
        /* �b���S���ϥγc��ӭn�D�A�]������ */
    }
    return 0;
}

int main(void)
{	
	uchar i = 0;
    wdt_enable(WDTO_1S);
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
	DDRD = 0xeb;    //    1      1      1      0      1      0      1      1    
					//   out    out    out    D-     out     D+     out    out
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
	PORTD = 0x00;   //    0      0      0      0      0      0      0      0  
					// ---------------------------------------------------------
	
    usbInit();
    usbDeviceDisconnect();  /* �j��s�˸m�C�|�A����榹�Ƶ{���ɡA���_�O�Q���઺ */
	
    while(--i){             /* ����USB���Y�ޥXPC�D���A���������ʧ@�A �ɶ� > 250 ms */
        wdt_reset();		/* ���m�ݪ����p�ɾ� */
        _delay_ms(1);
    }
    usbDeviceConnect();
    sei();					/* �P����餤�_ */
    while(1)				/* �D�n���ƥ�j�� */
	{                
        wdt_reset();		/* ���m�ݪ����p�ɾ� */
        usbPoll();			/* ����USB����ǿ� */
		SEG(SEG_Data[0],SEG_Data[1],SEG_Data[2],SEG_Data[3]);	/* ����SEG��T */
    }
    return 0;
}