/*   ���عϮ�: USB�����]�p�P���γ]�p�J������d�ҵ{��-CH12(�\�éM�s��)                                */
/*   �{���\��: �Q��PB��PD�X��8��LED�A�H�γz�LPC0~PC5�BPD1�PPD3�X��4��8x8�I�x�}�A�M�ϥ�PD0Ū�������  */   

#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* sei():�P�त�_ */
#include <util/delay.h>     /*  _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* usbdrv.h�һݪ��ѼƩw�q */
#include "usbdrv.h"
#include "oddebug.h"        /* �ϥΰ����������d�� */

static uchar keyin;
static uchar LED[4][8];
/* �P���PC�BPD1�PPD3��X����8*8�I�x�} */
static void outputByte(uchar b)	
{
    PORTC = (PORTC & ~0x3f) | (b      & 0x3f);
    PORTD = (PORTD & ~0x02) | (b >> 5 & 0x02);
	PORTD = (PORTD & ~0x08) | (b >> 4 & 0x08);
}
/* �P���PB�PPD��X����LED */
static void outputLED(uchar b)
{
	PORTB = (PORTB & ~0x3F) | (b & 0x3F);
	PORTD = (PORTD & ~0xC0) | (b & 0xC0);
}

// 74LS273 �ɯ�
static void outputControl(uchar b)
{
	PORTD = (PORTD & ~0x20) | (b & 0x20);
}
// �z�LPD0Ū�������
static uchar inputKeys()
{
	uchar b;
	b = PIND & 0x01;
	return b;
}
/* 8x8�I�x�}����禡 */
static void Matrix(uchar Row, uchar Data)
{
	uchar i;
	for(i = 0 ; i < 8 ; i++)
	{
		wdt_reset();
		usbPoll();
		keyin = inputKeys();
		outputByte(Row + i);
		PORTC = (PORTC & ~0x40) | (0x40);
		//_delay_ms(1);
		PORTC = (PORTC & ~0x40) | (0x00);		
		outputByte(LED[Data][i]);
		_delay_ms(1);
	}
}

/* ------------------------------------------------------------------------- */
/* ------------------------------ USB ���� --------------------------------- */
/* ------------------------------------------------------------------------- */

PROGMEM const char usbHidReportDescriptor[22] = {     /* USB ���i�y�z�� */
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

/* usbFunctionRead():��D���n�qUSB�˸m�n�D�j���ƮɴN�|�I�susbFunctionRead() */ 
/* �����T���i�H�i�H�Ѧ�usbdrv/usbdrv.h��� */
uchar   usbFunctionRead(uchar *data, uchar len)
{
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		*(data) = keyin;	// �����
		*(data+1) = 0;		// �O�d (0)
		*(data+2) = 0;		// �O�d (0)
		*(data+3) = 0;		// �O�d (0)
		*(data+4) = 0;		// �O�d (0)
		*(data+5) = 0;		// �O�d (0)
		*(data+6) = 0;		// �O�d (0)
		*(data+7) = 0;		// �O�d (0)
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
        return 1;               	/* USB�ǿ鵲�� */
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		
		if(*(data) == 0x20) {
			LED[data[1]][data[2]] = data[3];
		} 
		else if(*(data) == 0x21) 
		{
			outputLED(data[1]);
			outputControl(0x20);	// ---- ���ͤ@�Ӯɯ�
			outputControl(0x0);		
		}
	}
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0;		/* �p�G�o�O�̫�@���ơA�N�^��1 */
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
	DDRC = 0x7F;    //    0      1      1      1      1      1      1      1  
					//  (Read)  out    out    out    out    out    out    out							 
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
    usbDeviceDisconnect();  /* �j��s�˸m�C�|�A����榹�Ƶ{���ɡA���_�O�Q���઺ */
	i = 0;
    while(--i){             /* ����USB���Y�ޥXPC�D���A���������ʧ@�A �ɶ� > 250 ms */
        wdt_reset();		/* ���m�ݪ����p�ɾ� */
        _delay_ms(1);
    }
    usbDeviceConnect();
    sei();					/* �P����餤�_ */
    while(1){               /* �D�n���ƥ�j�� */ 
        wdt_reset();		/* ���m�ݪ����p�ɾ� */
        usbPoll();			/* ����USB����ǿ� */
		Matrix(0x08,0);		/* ����8x8�I�x�}U3-1��T */
		Matrix(0x10,1);		/* ����8x8�I�x�}U3-2��T */
		Matrix(0x20,2);		/* ����8x8�I�x�}U3-3��T */
		Matrix(0x40,3);		/* ����8x8�I�x�}U3-4��T */
    }
    return 0;
}