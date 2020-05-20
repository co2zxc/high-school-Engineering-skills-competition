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
uchar LOGO=0,z=0,program_run=0;
/*  �P���PB�PPD��X��74273(LED)�P74244(�I�x�})  */
static void outputByte(uchar b)
{
	PORTB = (PORTB & ~0x3F) | (b & 0x3F);
	PORTD = (PORTD & ~0xC0) | (b & 0xC0);
}
static void outputs(uchar b)
{
	PORTC = (PORTC & ~0x7F) | (b & 0x7F);
}
// 74273 �ɯ�
static void output273(uchar b)
{
	PORTD = (PORTD & ~0x20) | (b & 0x20);
}

// �z�LPD0Ū�������
static uchar inputKeys()
{
	uchar b;
	b = PIND & 0x03;
	return b;
}

// �z�LPD1Ū�������
/* 8x8�I�x�}����禡 */
static void Matrix(uchar Row, uchar Data)
{
	uchar i;
	for(i = 0 ; i < 8 ; i++)
	{
		wdt_reset();
		usbPoll();
		keyin = inputKeys();
		outputs(Row+i);	
		outputByte(LED[Data][i]);
		_delay_ms(1);
	}
		
}
static void Matrixbf(uchar Row, uchar Data)
{
	uchar num[20][8] = {{0x00,0xc2,0xa1,0x91,0x89,0x86,0x00,0x00},  		
					{0x41,0x81,0x89,0x95,0x63,0x00,0x00,0x66},			
					{0x66,0x00,0x00,0x4f,0x89,0x89,0x89,0x71},			
					{0x00,0x00,0x06,0x89,0x89,0x49,0x3e,0x00},			//23:59
					
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},			
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x66},			
					{0x66,0x00,0x00,0x38,0x24,0x22,0xff,0x20},			
					{0x00,0x00,0x01,0x01,0xf9,0x05,0x03,0x00},			//  :47
					
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},		
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},			
					{0x00,0x00,0x00,0x7e,0x81,0x81,0x81,0x7e},			
					{0x00,0x00,0x7e,0x81,0x81,0x81,0x7e,0x00},//   00
					
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},			
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x66},			
					{0x66,0x00,0x00,0x7e,0x81,0x81,0x81,0x7e},			
					{0x00,0x00,0x7e,0x81,0x81,0x81,0x7e,0x00},    // :00
					
					{0x00,0x7e,0x81,0x81,0x81,0x7e,0x00,0x00},
					{0x7e,0x81,0x81,0x81,0x7e,0x00,0x00,0x66},
					{0x66,0x00,0x00,0x7e,0x81,0x81,0x81,0x7e},
					{0x00,0x00,0x7e,0x81,0x81,0x81,0x7e,0x00}};   // 00:00
								
					
					
	uchar pic[64] = {0x3f,0xa1,0xe1,0xe1,0xe1,0xe1,0xa1,0x3f,
					0x00,0x06,0x05,0x09,0x91,0xe1,0xe1,0x91,
					0x09,0x05,0x06,0x00,0xc0,0xc0,0xf8,0x88,
					0x8f,0x81,0x8f,0x88,0xf8,0xc0,0xc0,0x00,
					0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
					0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
					0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
					0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00};
	uchar i;
	
	if(keyin==0x00)  
	{
		for(i = 0 ; i < 8 ; i++)
		{
			wdt_reset();
			usbPoll();
			keyin = inputKeys();
			outputs(Row+i);	
			outputByte(pic[Data*8+i+LOGO]);
			_delay_ms(1);
		}		
	}
	else
	{
		for(i = 0 ; i < 8 ; i++)
		{
			wdt_reset();
			usbPoll();
			keyin = inputKeys();
			outputs(Row+i);	
			outputByte(num[Data][i]);
			_delay_ms(1);
		}
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
		*(data) = keyin;		// k0 k1
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
		
		if(*(data) == 0x22) 
		{
			program_run = 1;
		}
		else if(*(data) == 0x20) {
			LED[data[1]][data[2]] = data[3];
		} 
		else if(*(data) == 0x21) 
		{
			outputByte(data[1]);
			output273(0x0);	// ---- ���ͤ@�Ӯɯ�
			output273(0x20);		
		}
		else if(*(data) == 0x23)
		{
			program_run =2;  //����
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
    //wdt_enable(WDTO_1S);
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
    PORTC = 0x00;   //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xE8;    //    1      1      1      0      1      0      0      0    
					//   out    out    out    D-     out    D+     in     in
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
    sei();						/* �P����餤�_ */
	i = 0;
	outputByte(0xff);
	output273(0x0);				// ---- ���ͤ@�Ӯɯ�
	output273(0x20);		
    while(1)					// �D�n���ƥ�j�� */ 
	{	
		if(program_run ==2 )  //00:00
		{
			if(keyin == 0x03){
			Matrixbf(0x40,16);		// ����8x8�I�x�}U3-1��T 
			Matrixbf(0x20,17);		// ����8x8�I�x�}U3-2��T 
			Matrixbf(0x10,18);		// ����8x8�I�x�}U3-3��T 
			Matrixbf(0x08,19);		// ����8x8�I�x�}U3-4��T
			}
			if(keyin == 0x01 )		//   00
			{
				Matrixbf(0x40,8);		// ����8x8�I�x�}U3-1��T 
				Matrixbf(0x20,9);		// ����8x8�I�x�}U3-2��T 
				Matrixbf(0x10,10);		// ����8x8�I�x�}U3-3��T 
				Matrixbf(0x08,11);		// ����8x8�I�x�}U3-4��T 
			}
			if(keyin == 0x02)		//  :00
			{
				Matrixbf(0x40,12);		// ����8x8�I�x�}U3-1��T 
				Matrixbf(0x20,13);		// ����8x8�I�x�}U3-2��T 
				Matrixbf(0x10,14);		// ����8x8�I�x�}U3-3��T 
				Matrixbf(0x08,15);		// ����8x8�I�x�}U3-4��T 	
			}
			if(keyin == 0x00)
			{
				Matrixbf(0x40,4);		// ����8x8�I�x�}U3-1��T 
				Matrixbf(0x20,4);		// ����8x8�I�x�}U3-2��T 
				Matrixbf(0x10,4);		// ����8x8�I�x�}U3-3��T 
				Matrixbf(0x08,4);// ����8x8�I�x�}U3-4��T 	
			}
		}
		if(program_run == 1)
		{
			wdt_reset();		// ���m�ݪ����p�ɾ� 
			usbPoll();			// ����USB����ǿ� 
			Matrix(0x40,0);		// ����8x8�I�x�}U3-1��T 
			Matrix(0x20,1);		// ����8x8�I�x�}U3-2��T 
			Matrix(0x10,2);		// ����8x8�I�x�}U3-3��T 
			Matrix(0x08,3);		// ����8x8�I�x�}U3-4��T 		

		}
		if(program_run == 0)
		{
			if(keyin == 0x00)	// LOGO
			{
				Matrixbf(0x40,0);		// ����8x8�I�x�}U3-1��T 
				Matrixbf(0x20,1);		// ����8x8�I�x�}U3-2��T 
				Matrixbf(0x10,2);		// ����8x8�I�x�}U3-3��T 
				Matrixbf(0x08,3);		// ����8x8�I�x�}U3-4��T 
				if(z==0)
				{
					LOGO++;
					_delay_ms(10);
					if(LOGO>31)	z=1;
				}
				if(z==1)
				{
					LOGO--;
					_delay_ms(10);
					if(LOGO<1)	z=0;
				}
			}
			if(keyin == 0x01)		//   00
			{
				Matrixbf(0x40,8);		// ����8x8�I�x�}U3-1��T 
				Matrixbf(0x20,9);		// ����8x8�I�x�}U3-2��T 
				Matrixbf(0x10,10);		// ����8x8�I�x�}U3-3��T 
				Matrixbf(0x08,11);		// ����8x8�I�x�}U3-4��T 
			}
			if(keyin == 0x02)		//  :47
			{
				Matrixbf(0x40,4);		// ����8x8�I�x�}U3-1��T 
				Matrixbf(0x20,5);		// ����8x8�I�x�}U3-2��T 
				Matrixbf(0x10,6);		// ����8x8�I�x�}U3-3��T 
				Matrixbf(0x08,7);		// ����8x8�I�x�}U3-4��T 	
			}
			if(keyin == 0x03)		//23:59
			{
				Matrixbf(0x40,0);		// ����8x8�I�x�}U3-1��T 
				Matrixbf(0x20,1);		// ����8x8�I�x�}U3-2��T 
				Matrixbf(0x10,2);		// ����8x8�I�x�}U3-3��T 
				Matrixbf(0x08,3);		// ����8x8�I�x�}U3-4��T 		
			}

		}
		
	}
		return 0;
}