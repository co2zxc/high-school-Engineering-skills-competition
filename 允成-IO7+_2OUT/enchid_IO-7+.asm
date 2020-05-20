;******************************************************
;
;File: enchid.asm
;adapted from Cypress Semiconductor's logo.asm by Jan Axelson
;Date: 07/1/2004
;Chip: Cypress Semiconductor CY7C637xx Encore series USB Microcontroller
;Assembler: cyasm.exe

;Purpose: demonstrates USB communications with a HID-class device
;Description:
;Handles standard USB and HID-class requests.

;Supports Input reports using control and interrupt transfers,
;Output reports using control transfers,
;and Feature IN and OUT reports.

;The Input report returns the data received in the last Output report.
;The IN Feature report returns the data received in the last OUT Feature report.

;I used Cypress Semiconductor's logo.asm example code as a base 
;in creating this program. Most of the code is Cypress'.

;Companion host software is at www.Lvr.com/hidpage.htm

;For more information, visit Lakeview Research at http://www.lvr.com .
;Send comments, bug reports, etc. to jan@lvr.com .

;7/1/2004
;Added support for Feature reports and Input reports using control transfers.

;2/19/2001
;Changed endpoint 1 ISR so it doesn't toggle the data toggle on NAK.

;	Target: Cypress CY7C63743
;
;	Overview:  There is only one task handled by this
;				firmware, and that is USB.
;
;	USB
;		At bus reset the USB interface is re-initialized,
;		and the firmware soft reboots.  We are then able to
;		handle the standard chapter nine requests on 
;		endpoint zero (the control endpoint).  After this
;		device enumerates as a HID mouse on the USB, the
;		requests come to endpoint one (the data endpoint).
;		Endpoint one is used to send Input reports.
;		Endpoint zero is used to receive Output reports
;		in Set_Report requests.
;
;	Pin Connections
;
;			 -------------------
;			| P0[0]		P0[4]	|
;			| P0[1]		P0[5]	|
;			| P0[2]		P0[6]	|
;			| P0[3] 	P0[7]	|
;			| P1[0]		P1[1]	|	
;			| P1[2]		P1[3]	|
;			| P1[4]		P1[5]	|
;			| P1[6]		P1[7]	|
;	GND		| VSS		D+/SCLK	|	USB D+ / PS2 SCLK
;	GND		| VPP		D-/SDATA|	USB D- / PS2 SDATA
;	PULLUP	| VREG		VCC		|	+5
;			| XTALIN	XTALOUT	|
;			 -------------------
;
; Revisions:
;			9-25-2000			Creation
;
;To do (things to be added):
;Add ability to send and receive Feature reports.
;Add ability to receive Output reports > 8 bytes.
;Add ability to receive interrupt Out reports.
;Move watchdog reset out of 1-ms timer routine.
;
;**********************************************************
;
;	Copyright 2000 Cypress Semiconductor and Lakeview Research
;	Portions of this code are provided by Cypress and by
;	Lakeview Research as a reference.  Cypress and Lakeview
;	make no claims or warranties to this firmware's 
;	suitability for any application.
;
;********************************************************** 

;**************** assembler directives ***************** 

	CPU	63743

	XPAGEON

	INCLUDE	"637xx.inc"
	INCLUDE "USB.inc"

ep1_dmabuff:			equ	F0h
ep1_dmabuff0:			equ	ep1_dmabuff+0
ep1_dmabuff1:			equ	ep1_dmabuff+1
ep1_dmabuff2:			equ	ep1_dmabuff+2
ep1_dmabuff3:			equ	ep1_dmabuff+3
ep1_dmabuff4:			equ	ep1_dmabuff+4
ep1_dmabuff5:			equ	ep1_dmabuff+5
ep1_dmabuff6:			equ	ep1_dmabuff+6
ep1_dmabuff7:			equ	ep1_dmabuff+7

ep0_dmabuff:			equ	F8h
ep0_dmabuff0:			equ	ep0_dmabuff+0
ep0_dmabuff1:			equ	ep0_dmabuff+1
ep0_dmabuff2:			equ	ep0_dmabuff+2
ep0_dmabuff3:			equ	ep0_dmabuff+3
ep0_dmabuff4:			equ	ep0_dmabuff+4
ep0_dmabuff5:			equ	ep0_dmabuff+5
ep0_dmabuff6:			equ	ep0_dmabuff+6
ep0_dmabuff7:			equ	ep0_dmabuff+7

bmRequestType:			equ	ep0_dmabuff0
bRequest:				equ	ep0_dmabuff1
wValuelo:				equ	ep0_dmabuff2
wValuehi:				equ	ep0_dmabuff3
wIndexlo:				equ	ep0_dmabuff4
wIndexhi:				equ	ep0_dmabuff5
wLengthlo:				equ	ep0_dmabuff6

wLengthhi:				equ	ep0_dmabuff7


; DATA MEMORY VARIABLES
;
suspend_count:			equ	20h		; counter for suspend/resume
ep1_data_toggle:		equ	21h		; data toggle for INs on endpoint one
ep0_data_toggle: 		equ	22h		; data toggle for INs on endpoint zero




data_start:				equ	23h		; address of request response data, as an offset
data_count:				equ	24h		; number of bytes to send back to the host
maximum_data_count:		equ	25h		; request response size
ep0_in_machine:			equ	26h		
ep0_in_flag:			equ	27h
configuration:			equ	28h
ep1_stall:				equ	29h
idle:					equ	2Ah
protocol:				equ	2Bh
temp:					equ	2Ch		; temporary register
event_machine:			equ	2Dh
pending_data:			equ	2Eh
int_temp:				equ	2Fh
idle_timer:				equ	30h
idle_prescaler:			equ	31h
logo_index:				equ	32h
ep0_transtype:			equ	33h

data_byte_0:			equ	34h
data_byte_1:			equ	35h
output_report_received:		equ	36h
output_or_feature:			equ	37h


feature_report_byte_0:			equ	38h
feature_report_byte_1:			equ	ep1_dmabuff+1
feature_report_byte_2:			equ	ep1_dmabuff+2
feature_report_byte_3:			equ	ep1_dmabuff+3
feature_report_byte_4:			equ	ep1_dmabuff+4
feature_report_byte_5:			equ	ep1_dmabuff+5
feature_report_byte_6:			equ	ep1_dmabuff+6
feature_report_byte_7:			equ	ep1_dmabuff+7

input_report_byte_0:			equ	40h
input_report_byte_1:			equ	ep1_dmabuff+1
input_report_byte_2:			equ	ep1_dmabuff+2
input_report_byte_3:			equ	ep1_dmabuff+3
input_report_byte_4:			equ	ep1_dmabuff+4
input_report_byte_5:			equ	ep1_dmabuff+5
input_report_byte_6:			equ	ep1_dmabuff+6
input_report_byte_7:			equ	ep1_dmabuff+7

output_report_byte_0:			equ	48h
output_report_byte_1:			equ	ep1_dmabuff+1
output_report_byte_2:			equ	ep1_dmabuff+2
output_report_byte_3:			equ	ep1_dmabuff+3
output_report_byte_4:			equ	ep1_dmabuff+4
output_report_byte_5:			equ	ep1_dmabuff+5
output_report_byte_6:			equ	ep1_dmabuff+6
output_report_byte_7:			equ	ep1_dmabuff+7

port0_data_temp			:     equ 50h
port1_data_temp			:     equ 51h


; STATE MACHINE CONSTANTS
;EP0 IN TRANSACTIONS
EP0_IN_IDLE:			equ	00h
CONTROL_READ_DATA:		equ	02h
NO_DATA_STATUS:			equ	04h
EP0_IN_STALL:			equ	06h

; FLAG CONSTANTS
;EP0 NO-DATA CONTROL FLAGS
ADDRESS_CHANGE_PENDING:	equ	00h
NO_CHANGE_PENDING:		equ	02h

; RESPONSE SIZES
DEVICE_STATUS_LENGTH:		equ	2
DEVICE_CONFIG_LENGTH:		equ	1
ENDPOINT_STALL_LENGTH:		equ 2
INTERFACE_STATUS_LENGTH:	equ 2
INTERFACE_ALTERNATE_LENGTH:	equ	1
INTERFACE_PROTOCOL_LENGTH:	equ	1

NO_EVENT_PENDING:			equ	00h
EVENT_PENDING:				equ	02h

;***** TRANSACTION TYPES

TRANS_NONE:						equ	00h
TRANS_CONTROL_READ:				equ	02h
TRANS_CONTROL_WRITE:			equ	04h
TRANS_NO_DATA_CONTROL:			equ	06h

;Additional notes:
;ep0_mode is the Endpoint 0 mode register (12h).


;*************** interrupt vector table ****************

ORG 00h			

jmp	reset				; reset vector		

jmp	bus_reset			; bus reset interrupt

jmp	error				; 128us interrupt

jmp	1ms_timer			; 1.024ms interrupt

jmp	endpoint0			; endpoint 0 interrupt

jmp	endpoint1			; endpoint 1 interrupt

jmp	error				; endpoint 2 interrupt

jmp	error				; reserved

jmp	error				; Capture timer A interrupt Vector

jmp	error				; Capture timer B interrupt Vector

jmp	error				; GPIO interrupt vector

jmp	error				; Wake-up interrupt vector 


;************** program listing ************************

ORG  1Ah
error: halt

;*******************************************************
;
;	Interrupt handler: reset
;	Purpose: The program jumps to this routine when
;		 the microcontroller has a power on reset.
;
;*******************************************************

reset:
	; set for use with external oscillator
	mov		A, (LVR_ENABLE | EXTERNAL_CLK)	
	iowr	clock_config

	; setup data memory stack pointer
	mov		A, 68h
	swap	A, dsp		

	; clear variables
	mov		A, 00h		
	mov		[ep0_in_machine], A		; clear ep0 state machine
	mov		[configuration], A
	mov		[ep1_stall], A
	mov		[idle], A
	mov		[suspend_count], A
	mov		[ep1_dmabuff0], A
	mov		[ep1_dmabuff1], A
	mov		[ep1_dmabuff2], A
	mov		[int_temp], A
	mov		[idle_timer], A
	mov		[idle_prescaler], A
	mov		[event_machine], A
	mov		[logo_index], A
	mov		[ep0_transtype], A

	mov		A, 01h
	mov		[protocol], A

	;-------------------------------------------------------
	; Set Port0 : M1,M0 = 01
	;-------------------------------------------------------
	mov	A, FFh
	iowr	port0_mode0
	mov	A, 00h
	iowr	port0_mode1
	iowr	port0
	;-------------------------------------------------------
	; Set Port1 : M1,M0 = 01
	;-------------------------------------------------------
	mov	A, FFh
	iowr	port1_mode0
	mov	A, 00h
	iowr	port1_mode1
	iowr	port1




	; enable global interrupts
	mov		A, (1MS_INT | USB_RESET_INT)
	iowr 	global_int

	; enable endpoint  0 interrupt
	mov 	A, EP0_INT			
	iowr 	endpoint_int

	; enable USB address for endpoint 0
	mov		A, ADDRESS_ENABLE
	iowr	usb_address

	; enable all interrupts
	ei

	; enable USB pullup resistor
	mov		A, VREG_ENABLE  
	iowr	usb_status

task_loop:


;-------------------------------------------------------
	; Set Port0 : M1,M0 = 01
	;-------------------------------------------------------
	mov	A, FFh
	iowr	port0_mode0
	mov	A, 00h
	iowr	port0_mode1
;	iowr	port0

			mov		A, [port0_data_temp]
			iowr		port0

	;-------------------------------------------------------
	; Set Port1 : M1,M0 = 01
	;-------------------------------------------------------
	mov	A, FFh
	iowr	port1_mode0
	mov	A, 00h
	iowr	port1_mode1
;	iowr	port1

			mov		A, [port1_data_temp]
			iowr		port1
;--------------------------------------------------------------------
	mov		A, [event_machine]

	jacc	event_machine_jumptable
		no_event_pending:

		mov 	A, [output_report_received]
		cmp	A, 0

		jz 	no_input_report_to_send

			; if not configured then skip data transfer
			mov		A, [configuration]
			cmp		A, 01h
			jnz		no_event_task
			; if stalled then skip data transfer
			mov		A, [ep1_stall]
			cmp		A, FFh
			jz		no_event_task
			
			;copy the received output report data to endpoint 1's buffer
			;to return to the host in an Input report.
			mov		A, [output_report_byte_0]
			mov		[ep1_dmabuff0], A
			mov		A, [output_report_byte_0]
			mov		[ep1_dmabuff0], A

			mov		A, 02h		; set endpoint 1 to send 2 bytes
			or		A, [ep1_data_toggle]
			iowr	ep1_count
			mov		A, ACK_IN		; set to ack on endpoint 1
			iowr	ep1_mode	

;Set output_report_received to 0 to indicate the report has been handled.
			mov		A, 00h
			mov		[output_report_received], A

clear_pending_events:
			mov		A, EVENT_PENDING	; clear pending events
			mov		[event_machine], A 




jmp task_loop
		no_input_report_to_send:
			mov		A, NAK_IN_OUT		; set to NAK on endpoint 1
			iowr	ep1_mode	

		event_task_done:



	no_event_task:
	event_pending:




	jmp task_loop


;*******************************************************
;
;	Interrupt handler: bus_reset
;	Purpose: The program jumps to this routine when
;		 the microcontroller has a bus reset.
;
;*******************************************************

bus_reset:
	mov		A, STALL_IN_OUT			; set to STALL INs&OUTs
	iowr	ep0_mode

	mov		A, ADDRESS_ENABLE		; enable USB address 0
	iowr	usb_address
	mov		A, DISABLE				; disable endpoint1
	iowr	ep1_mode

	mov		A, 00h					; reset program stack pointer
	mov		psp,a	

	jmp		reset


;*******************************************************
;
;	Interrupt: 1ms_clear_control
;	Purpose: Every 1ms this interrupt handler clears
;		the watchdog timer.
;
;*******************************************************

1ms_timer:
	push A

	; clear watchdog timer
	iowr watchdog

	; check for no bus activity/usb suspend
  1ms_suspend_timer:
	iord	usb_status				; read bus activity bit
	and		A, BUS_ACTIVITY			; mask off activity bit

	jnz		bus_activity
			
	inc		[suspend_count]			; increment suspend counter
	mov		A, [suspend_count]
	cmp		A, 04h					; if no bus activity for 3-4ms,
	jz		usb_suspend				; then go into low power suspend
	jmp		ms_timer_done

 	usb_suspend:

		; enable wakeup timer

		mov		A, (USB_RESET_INT)
		iowr 	global_int

  	iord	control
  	or		A, SUSPEND			; set suspend bit
 		ei
 		iowr	control
 		nop

		; look for bus activity, if none go back into suspend
		iord	usb_status
		and		A, BUS_ACTIVITY
		jz		usb_suspend		

		; re-enable interrupts
		mov		A, (1MS_INT | USB_RESET_INT)
		iowr 	global_int


	bus_activity:
		mov		A, 00h				; reset suspend counter
		mov		[suspend_count], A
		iord	usb_status
		and		A, ~BUS_ACTIVITY	; clear bus activity bit
		iowr	usb_status


	ms_timer_done:
		pop A
		reti

;*******************************************************
;
;	Interrupt: endpoint0
;	Purpose: Usb control endpoint handler.  This interrupt
;			handler formulates responses to SETUP and 
;			CONTROL READ, and NO-DATA CONTROL transactions. 
;
;	Jump table entry formulation for bmRequestType and bRequest
;
;	1. Add high and low nibbles of bmRequestType.
;	2. Put result into high nibble of A.
;	3. Mask off bits [6:4].
;	4. Add bRequest to A.
;	5. Double value of A (jmp is two bytes).
;
;*******************************************************

endpoint0:
	push	X
	push	A

;Reading ep0_mode enables writing to the endpoint's buffer.
	iord	ep0_mode
;If EP0_ACK isn't set, the transaction didn't complete with an Ack,
;so exit the ISR.
	and		A, EP0_ACK
	jz		ep0_done

;Bit 5, 6, or 7 is set to indicate whether the transaction is
;Setup, In, or Out. Find out which it is and jump to a routine to handle it.
	iord	ep0_mode
	asl		A

	jc		ep0_setup_received
	asl		A
	jc		ep0_in_received
	asl		A
	jc		ep0_out_received

  ep0_done:
	pop		A
	pop		X
	reti

	ep0_setup_received:

 		mov		A, NAK_IN_OUT				; clear setup bit to enable
		iowr	ep0_mode					; writes to EP0 DMA buffer




		mov		A, [bmRequestType]			; compact bmRequestType into 5 bit field
		and		A, E3h						; clear bits 4-3-2, these unused for our purposes
		push	A							; store value
		asr		A							; move bits 7-6-5 into 4-3-2's place
		asr		A
		asr		A
		mov		[int_temp], A				; store shifted value
		pop		A							; get original value
		or		A, [int_temp]				; or the two to get the 5-bit field
		and		A, 1Fh						; clear bits 7-6-5 (asr wraps bit7)
		asl		A							; shift to index jumptable

;Use bmRequestType to find out whether the request is host-to-device (h2d)
;or device-to-host (d2h), and whether the request is for the device, an ;interface, or an endpoint.
		jacc	bmRequestType_jumptable		; jump to handle bmRequestType

;Use the request's value to decide where to jump.
;Shift left to double the request's value because the entries in the jumptable
;are two bytes each.

		h2d_std_device:
			mov		A, [bRequest]
			asl		A
			jacc	h2d_std_device_jumptable

		h2d_std_interface:	
			mov		A, [bRequest]
			asl		A
			jacc	h2d_std_interface_jumptable

		h2d_std_endpoint:
			mov		A, [bRequest]
			asl		A
			jacc	h2d_std_endpoint_jumptable

;I added this jump for set_report requests
		h2d_class_interface:
			mov		A, [bRequest]
			asl		A
			jacc	h2d_class_interface_jumptable

		d2h_std_device:
			mov		A, [bRequest]
			asl		A
			jacc	d2h_std_device_jumptable

		d2h_std_interface:
			mov		A, [bRequest]
			asl		A
			jacc	d2h_std_interface_jumptable

		d2h_std_endpoint:
			mov		A, [bRequest]
			asl		A
			jacc	d2h_std_endpoint_jumptable

;I added this jump for get_report requests
		d2h_class_interface:
			mov		A, [bRequest]
			asl		A
			jacc	d2h_class_interface_jumptable

	;;************ DEVICE REQUESTS **************

	set_device_address:						; SET ADDRESS
		mov		A, ADDRESS_CHANGE_PENDING	; set flag to indicate we
		mov		[ep0_in_flag], A			; need to change address on
		mov		A, [wValuelo]
		mov		[pending_data], A
		jmp		initialize_no_data_control

	set_device_configuration:				; SET CONFIGURATION
		mov		A, [wValuelo]
		cmp		A, 01h
		jz		configure_device
;0 means unconfigure the device. 
		unconfigure_device:					; set device as unconfigured
			mov		[configuration], A
			mov		A, DISABLE				; disable endpoint 1
			iowr	ep1_mode
			mov 	A, EP0_INT				; turn off endpoint 1 interrupts
			iowr 	endpoint_int
			jmp		set_device_configuration_done
;1 means select configuration 1 (the only configuration option).
		configure_device:					; set device as configured
			mov		[configuration], A

			mov		A, [ep1_stall]			; if endpoint 1 is stalled
			and		A, FFh
			jz		ep1_nak_in_out
				mov		A, STALL_IN_OUT		; set endpoint 1 mode to stall
				iowr	ep1_mode
				jmp		ep1_set_int
			ep1_nak_in_out:
				mov		A, NAK_IN_OUT		; otherwise set it to NAK in/out
				iowr	ep1_mode
			ep1_set_int:
			mov 	A, EP0_INT | EP1_INT 	; enable endpoint 1 interrupt		
			iowr 	endpoint_int
			mov		A, 00h
			mov		[ep1_data_toggle], A	; reset the data toggle
			mov		[ep1_dmabuff0], A		; reset endpoint 1 fifo values
			mov		[ep1_dmabuff1], A
			mov		[ep1_dmabuff2], A
			set_device_configuration_done:
			mov		A, NO_CHANGE_PENDING
			mov		[ep0_in_flag], A
			jmp		 initialize_no_data_control


	get_device_status:						; GET STATUS
		mov		A, DEVICE_STATUS_LENGTH
		mov		[maximum_data_count], A
		mov		A, (device_status_wakeup_disabled - control_read_table)
		mov		[data_start], A
		jmp		initialize_control_read
		

	get_device_descriptor:					; GET DESCRIPTOR
		mov		A, [wValuehi]

		asl		A
		jacc	get_device_descriptor_jumptable

		send_device_descriptor:
			mov		A, 00h					; get device descriptor length
			index	device_desc_table
			mov		[maximum_data_count], A
			mov		A, (device_desc_table - control_read_table)
			mov		[data_start], A
			jmp		initialize_control_read

		send_configuration_descriptor:
			mov		A, 02h
   			index	config_desc_table:
			mov		[maximum_data_count], A
			mov		A, (config_desc_table - control_read_table)
			mov		[data_start], A
			jmp		initialize_control_read

		send_string_descriptor:
			mov		A, [wValuelo]
			asl		A
			jacc	string_jumptable:

			language_string:
				mov		A, 00h
				index	ilanguage_string
				mov		[maximum_data_count], A
				mov		A, (ilanguage_string - control_read_table)
				mov		[data_start], A
				jmp		initialize_control_read

			manufacturer_string:
				mov		A, 00h
				index	imanufacturer_string
				mov		[maximum_data_count], A
				mov		A, (imanufacturer_string - control_read_table)
				mov		[data_start], A
				jmp		initialize_control_read

			product_string:
				mov		A, 00h
				index	iproduct_string
				mov		[maximum_data_count], A
				mov		A, (iproduct_string - control_read_table)
				mov		[data_start], A
				jmp		initialize_control_read

		send_interface_descriptor:
			mov		A, 00h					; get interface descriptor length
			index	interface_desc_table 
			mov		[maximum_data_count], A
			mov		A, (interface_desc_table - control_read_table)
			mov		[data_start], A
			jmp		initialize_control_read

		send_endpoint_descriptor:
			mov		A, 00h					; get endpoint descriptor length
			index	endpoint_desc_table
			mov		[maximum_data_count], A
			mov		A, (endpoint_desc_table - control_read_table)
			mov		[data_start], A
			jmp		initialize_control_read


		get_device_configuration:			; GET CONFIGURATION
			mov		A, DEVICE_CONFIG_LENGTH
			mov		[maximum_data_count], A

			mov		A, [configuration]		; test configuration status
			and		A, FFh
			jz		device_unconfigured
			device_configured:				; send configured status
				mov		A, (device_configured_table - control_read_table)
				mov		[data_start], A
				jmp		initialize_control_read
			device_unconfigured:				; send unconfigured status
				mov		A, (device_unconfigured_table - control_read_table)
				mov		[data_start], A
				jmp		initialize_control_read


	;;************ INTERFACE REQUESTS ***********

	set_interface_interface:				; SET INTERFACE
		mov		A, [wValuelo]
		cmp		A, 00h						; there are no alternate interfaces
		jz		alternate_supported			; for this device
		alternate_not_supported:			; if the host requests any other
			jmp		request_not_supported	; alternate than 0, stall.	
		alternate_supported:
			mov		A, NO_CHANGE_PENDING
			mov		[ep0_in_flag], A
			jmp		initialize_no_data_control


	get_interface_status:					; GET STATUS
		mov		A, INTERFACE_STATUS_LENGTH
		mov		[maximum_data_count], A

		mov		A, (interface_status_table - control_read_table)
		mov		[data_start], A
		jmp		initialize_control_read
		

	get_interface_interface:				; GET INTERFACE
		mov		A, INTERFACE_ALTERNATE_LENGTH
		mov		[maximum_data_count], A
		mov		A, (interface_alternate_table - control_read_table)
		mov		[data_start], A
		jmp		initialize_control_read


	set_interface_idle:						; SET IDLE
		mov		A, [wValuehi]				; test if new idle time 
		cmp		A, 00h						; disables idle timer

		jz		idle_timer_disable

		mov		A, [idle_timer]				; test if less than 4ms left
		cmp		A, 01h
		jz		set_idle_last_not_expired

		mov		A, [wValuehi]				; test if time left less than
		sub		A, [idle_timer]				; new idle value
		jnc		set_idle_new_timer_less

		jmp		set_idle_normal

		idle_timer_disable:
			mov		[idle], A				; disable idle timer
			jmp		set_idle_done

		set_idle_last_not_expired:
			mov		A, EVENT_PENDING		; send report immediately
			mov		[event_machine], A
			mov		A, 00h					; reset idle prescaler
			mov		[idle_prescaler], A
			mov		A, [wValuehi]			; set new idle value
			mov		[idle_timer], A
			mov		[idle], A
			jmp		set_idle_done

		set_idle_new_timer_less:			
			mov		A, 00h
			mov		[idle_prescaler], A		; reset idle prescaler
			mov		A, [wValuehi]
			mov		[idle_timer], A			; update idle time value
			mov		[idle], A
			jmp		set_idle_done

		set_idle_normal:
			mov		A, 00h					; reset idle prescaler
			mov		[idle_prescaler], A
			mov		A, [wValuehi]			; update idle time value
			mov		[idle_timer], A
			mov		[idle], A

		set_idle_done:
			mov		A, NO_CHANGE_PENDING	; respond with no-data control
			mov		[ep0_in_flag], A		; transaction
			jmp		initialize_no_data_control


	set_interface_protocol:					; SET PROTOCOL
		mov		A, [wValuelo]
		mov		[protocol], A				; set protocol value
		mov		A, NO_CHANGE_PENDING
		mov		[ep0_in_flag], A			; respond with no-data control
		jmp		initialize_no_data_control	; transaction


	get_report:					; GET REPORT

		mov		A, TRANS_CONTROL_READ		; set transaction type to control read
		mov		[ep0_transtype], A

		mov		A, DATA_TOGGLE				; set data toggle to DATA ONE
		mov		[ep0_data_toggle], A
;Send NAK in response to any new INs or OUTs while writing to the buffers.
		mov		A, NAK_IN_OUT				; clear setup bit to write to
		iowr	ep0_mode					; endpoint fifo

;Get the report type (Input or Feature) from the low byte of the Value field.
		mov A, [wValuehi]
		;Double the value for the jump table
		asl A
		jacc in_report_type_jumptable

send_input_report:


;If the host requests the Input report with a control transfer, 
;copy the Input report data from the Input report buffer to Endpoint 0.
		mov		A, [output_report_byte_0]
		mov		[ep0_dmabuff0], A

		mov		A, [output_report_byte_1]
		mov		[ep0_dmabuff1], A

		mov		A, CONTROL_READ_DATA		; set state machine state
		mov		[ep0_in_machine], A			
		mov		X, 02h						; set number of bytes to transfer to 2
		jmp		dmabuffer_load_done			; jump to finish transfer
		
	
send_feature_report:

;If the host requests a Feature report, 
;copy the Feature report data from the Feature report buffer to Endpoint 0.
		mov		A, [feature_report_byte_0]
		mov		[ep0_dmabuff0], A

		mov		A, [feature_report_byte_1]
		mov		[ep0_dmabuff1], A

		mov		A, CONTROL_READ_DATA		; set state machine state
		mov		[ep0_in_machine], A			
		mov		X, 02h						; set number of bytes to transfer to 2
		jmp		dmabuffer_load_done			; jump to finish transfer



	get_interface_idle:						; GET IDLE
		mov		A, DATA_TOGGLE				; set data toggle to DATA ONE
		mov		[ep0_data_toggle], A
		mov		A, NAK_IN_OUT				; clear setup bit to write to
		iowr	ep0_mode					; endpoint fifo

		mov		A, [idle]					; copy over idle time
		mov		[ep0_dmabuff0], A

		mov		A, CONTROL_READ_DATA		; set state machine state
		mov		[ep0_in_machine], A			
		mov		X, 01h						; set number of byte to transfer to 3
		jmp		dmabuffer_load_done			; jump to finish transfer

	
	get_interface_protocol:					; GET PROTOCOL
		mov		A, INTERFACE_PROTOCOL_LENGTH
		mov		[maximum_data_count], A		; get offset of device descriptor table
		mov		A, [protocol]
		and		A, 01h
		jz		boot_protocol
		report_protocol:
			mov		A, (interface_report_protocol - control_read_table)
			mov		[data_start], A
			jmp		initialize_control_read	; get ready to send data
		boot_protocol:
			mov		A, (interface_boot_protocol - control_read_table)
			mov		[data_start], A
			jmp		initialize_control_read	; get ready to send data


	get_interface_hid:
		mov		A, [wValuehi]
		cmp		A, 21h
		jz		get_interface_hid_descriptor
		cmp		A, 22h
		jz		get_interface_hid_report
		jmp		request_not_supported

	get_interface_hid_descriptor:			; GET HID CLASS DESCRIPTOR
		mov		A, 00h						; get hid decriptor length
		index	hid_desc_table
		mov		[maximum_data_count], A		; get offset of device descriptor table
		mov		A, (hid_desc_table - control_read_table)
		mov		[data_start], A
		jmp		initialize_control_read		; get ready to send data


	get_interface_hid_report:				; GET HID REPORT DESCRIPTOR
		mov		A, 07h						; get hid report descriptor length
		index	hid_desc_table
		mov		[maximum_data_count], A		; get offset of device descriptor table
		mov		A, (hid_report_desc_table - control_read_table)
		mov		[data_start], A
		jmp		initialize_control_read		; get ready to send data


	;;************ ENDPOINT REQUESTS ************

	clear_endpoint_feature:					; CLEAR FEATURE
		mov		A, [wValuelo]
		cmp		A, ENDPOINT_STALL
		jnz		request_not_supported		
		mov		A, 00h						; clear endpoint 1 stall
		mov		[ep1_stall], A
		mov		A, NO_CHANGE_PENDING		; respond with no-data control
		mov		[ep0_in_flag], A
		jmp		initialize_no_data_control


	set_endpoint_feature:					; SET FEATURE
		mov		A, [wValuelo]
		cmp		A, ENDPOINT_STALL
		jnz		request_not_supported		
		mov		A, FFh						; stall endpoint 1
		mov		[ep1_stall], A
		mov		A, NO_CHANGE_PENDING		; respond with no-data control
		mov		[ep0_in_flag], A
		jmp		initialize_no_data_control

	get_endpoint_status:					; GET STATUS
		mov		A, ENDPOINT_STALL_LENGTH
		mov		[maximum_data_count], A
		mov		A, [ep1_stall]				; test if endpoint 1 stalled
		and		A, FFh
		jnz		endpoint_stalled
		endpoint_not_stalled:				; send no-stall status
			mov		A, (endpoint_nostall_table - control_read_table)
			mov		[data_start], A
			jmp		initialize_control_read
		endpoint_stalled:					; send stall status
			mov		A, (endpoint_stall_table - control_read_table)
			mov		[data_start], A
			jmp		initialize_control_read
		

	;;************ CLASS INTERFACE REQUESTS ************

	set_report:							; SET REPORT
;Find out how many bytes to read in the Out transaction(s) that will follow.
;This value is in WLengthlo and WLengthhi.
;Save the length in data_count.
		mov A, [wLengthlo]
		mov [data_count], A
;The Output reports are 2 bytes, so wLengthhi is unused.
		mov A, 0
		mov [wLengthhi], A

;Unlock the counter register so it can be updated with the 
;number of bytes in the data stage
		iord	ep0_count

;Get the report type (Output or Feature) from the high byte of the Value field.
		mov A, [wValuehi]
		;Double the value for the jump table
		asl A
		jacc out_report_type_jumptable

get_output_report:

;Receive an Output report from the host.

;output_or_feature indicates the report type.
		mov A, 0
		mov [output_or_feature], A



;Enable receiving data in an Out transaction.
		jmp initialize_control_write

get_feature_report:

;receive a Feature report from the host.



;output_or_feature indicates the report type.
		mov A, 1
		mov [output_or_feature], A

;Enable receiving data in an Out transaction.
		jmp initialize_control_write
;my code **************************************************

;;***************** CONTROL READ TRANSACTION **************

	initialize_control_read:
		mov		A, TRANS_CONTROL_READ		; set transaction type to control read
		mov		[ep0_transtype], A

		mov		A, DATA_TOGGLE				; set data toggle to DATA ONE
		mov		[ep0_data_toggle], A

		; if wLengthhi == 0
		mov		A, [wLengthhi]				; find lesser of requested and maximum
		cmp		A, 00h
		jnz		initialize_control_read_done
		; and wLengthlo < maximum_data_count
		mov		A, [wLengthlo]				; find lesser of requested and maximum
		cmp		A, [maximum_data_count]		; response lengths
		jnc		initialize_control_read_done
		; then maximum_data_count >= wLengthlo
		mov		A, [wLengthlo]
		mov		[maximum_data_count], A
		initialize_control_read_done:
			jmp		control_read_data_stage	; send first packet


;;***************** CONTROL WRITE TRANSACTION *************

	initialize_control_write:
		mov		A, TRANS_CONTROL_WRITE		; set transaction type to control write
;The firmware uses the value in ep0_transtype to decide how to respond
;to a token packet.
		mov		[ep0_transtype], A

		mov		A, DATA_TOGGLE				; set accepted data toggle
		mov		[ep0_data_toggle], A

;Send ACK in response to OUT packets, 
;which will contain the Control Write data.
;Send NAK in response to IN packets (not expected).
		mov		A, ACK_OUT_NAK_IN			; set mode
		iowr	ep0_mode

;Return from the Endpoint 0 ISR.
		pop		A
		pop		X
		reti

;;***************** NO DATA CONTROL TRANSACTION ***********

	initialize_no_data_control:
		mov		A, TRANS_NO_DATA_CONTROL	; set transaction type to no data control
		mov		[ep0_transtype], A

		mov		A, STATUS_IN_ONLY			; set SIE for STATUS IN mode
		iowr	ep0_mode
		pop		A
		pop		X
		reti


;;***************** UNSUPPORTED TRANSACTION ***************

	request_not_supported:
		iord	ep0_mode
		mov	A, STALL_IN_OUT					; send a stall to indicate that the request
		iowr	ep0_mode					; is not supported
		pop	A
		pop	X
		reti


;**********************************************************

	;**********************************
	; IN - CONTROL READ DATA STAGE
	;	 - CONTROL WRITE STATUS STAGE
	;	 - NO DATA CONTROL STATUS STAGE

	ep0_in_received:
	mov		A, [ep0_transtype]
	jacc	ep0_in_jumptable


	;**********************************

	control_read_data_stage:

		mov		X, 00h

		mov		A, [maximum_data_count]
		cmp		A, 00h						; has all been sent
		jz		dmabuffer_load_done			

		dmabuffer_load:
			mov		A, X					; check if 8 byte ep0 dma
			cmp		A, 08h					; buffer is full
			jz		dmabuffer_load_done
			mov		A, [data_start]			; read data from desc. table
			index	control_read_table
			mov		[X + ep0_dmabuff0], A

			inc		X						; increment buffer offset
			inc		[data_start]			; increment descriptor table pointer
			dec		[maximum_data_count]	; decrement number of bytes requested
			jz		dmabuffer_load_done
			jmp		dmabuffer_load			; loop to load more data
			dmabuffer_load_done:
	
		iord	ep0_count					; unlock counter register
		mov		A, X						; find number of bytes loaded
		or		A, [ep0_data_toggle]		; or data toggle
		iowr	ep0_count					; write ep0 count register

		mov		A, ACK_IN_STATUS_OUT		; set endpoint mode to ack next IN
		iowr	ep0_mode					; or STATUS OUT
			
		mov		A, DATA_TOGGLE				; toggle data toggle
		xor		[ep0_data_toggle], A

		pop		A
		pop		X
		reti


	;**********************************

	control_write_status_stage:
;Jump here if the device has received an IN token packet 
;with ep0_transtype = TRANS_CONTROL_WRITE
;The device has sent a 0-byte In data packet to complete the transfer
;because ep0_mode was set to Status_In_Only at the end of the data stage. 

		mov		A, STATUS_OUT_ONLY
		iowr	ep0_mode

	mov		A, TRANS_NONE
	mov		[ep0_transtype], A

		pop		A
		pop		X

		reti


	;**********************************

	no_data_control_status_stage:
		mov		A, [ep0_in_flag]		; end of no-data control transaction 
		cmp		A, ADDRESS_CHANGE_PENDING
		jnz		no_data_status_done

		change_address:
			mov		A, [pending_data]	; change the device address if this
			or		A, ADDRESS_ENABLE	; data is pending
			iowr	usb_address


		no_data_status_done:			; otherwise set to stall in/out until
			mov		A, STALL_IN_OUT		; a new setup
			iowr	ep0_mode

			mov		A, TRANS_NONE
			mov		[ep0_transtype], A

			pop		A
			pop		X
			reti


;**********************************************************

	;**********************************
	; OUT - CONTROL READ STATUS STAGE
	;	  - CONTROL WRITE DATA STAGE
	;	  - ERROR DURING NO DATA CONTROL TRANSACTION

	ep0_out_received:

;ep0_transtype was set during a previous transaction in the transfer.
	mov		A, [ep0_transtype]
	jacc	ep0_out_jumptable

	;**********************************

	control_read_status_stage:

		mov		A, STATUS_IN_ONLY
		iowr	ep0_mode

		mov		A, TRANS_NONE
		mov		[ep0_transtype], A

		pop		A
		pop		X
		reti	


	;**********************************

	control_write_data_stage:
;If the data-valid bit isn't set, we're done with the data stage.
;(logo.asm used mov A, ep0_count (incorrect) instead of iord here.)
		iord		ep0_count					; check that data is valid
		and		A, DATA_VALID
		jz		control_write_data_stage_done

;(logo.asm used mov A, ep0_count (incorrect) instead of iord here.)
		iord		ep0_count					;check data toggle
		and		A, DATA_TOGGLE
		xor		A, [ep0_data_toggle]
		jnz		control_write_data_stage_done

;The only supported control_write transfer is set_report.

;Is it an Output or Feature report?
			mov		A, 0
			cmp		A, [output_or_feature]
			jz		output_report

feature_report:

			mov		A, [ep0_dmabuff0]
			mov		[feature_report_byte_0], A
			mov		A, [ep0_dmabuff1]
			mov		[feature_report_byte_1], A

			jmp		data_toggle
output_report:

;Copy the report's two bytes to the data memory.

			mov		A, [ep0_dmabuff0]
			mov		[port0_data_temp]     , A
			mov		[output_report_byte_0], A
			iowr		port0
;			mov		[input_report_byte_0], A
			mov		A, [ep0_dmabuff1]
			mov		[port1_data_temp]     , A
			mov		[output_report_byte_1], A
			iowr		port1
;			mov		[input_report_byte_1], A





;Set output_report_received to 1 to indicate there's a new report.
			mov		A, 01h
			mov		[output_report_received], A








data_toggle:

;Toggle the data-toggle bit.
		mov		A, DATA_TOGGLE
		xor		[ep0_data_toggle], A

;If all of the data has been received,
;configure Endpoint 0 to send a 0-byte data packet in response
;to an IN packet (the transfer's status phase)
;and to Stall an Out packet.
;To do: add the ability to use > 1 data transaction.

		mov		A, STATUS_IN_ONLY
		iowr		ep0_mode

		control_write_data_stage_done:


		pop		A
		pop		X
		reti

		
	;**********************************

	no_data_control_error:	
		mov		A, STALL_IN_OUT
		iowr	ep0_mode

		mov		A, TRANS_NONE
		mov		[ep0_transtype], A

		pop		A
		pop		X
		reti


	
;*******************************************************
;
;	Interrupt handler: endpoint1
;	Purpose: This interrupt routine handles the specially
;		 reserved data endpoint 1.  This
;		 interrupt happens every time a host sends an
;		 IN on endpoint 1.  The data to send (NAK or 3
;		 byte packet) is already loaded, so this routine
;		 just prepares the dma buffers for the next packet
;
;*******************************************************

endpoint1:
	push	A
	iord	ep1_mode
; If the interrupt was due to sending a NAK, 
; exit the ISR.
	and	A, EP_ACK
	jz	     endpoint1_NAK

; Otherwise, toggle the data toggle for the next transaction.
	mov		A, 80h
	xor		[ep1_data_toggle], A

	mov		A, NO_EVENT_PENDING				; clear pending events
	mov		[event_machine], A 

	; set response
	mov		A, [ep1_stall]					; if endpoint is set to stall, then set
	cmp		A, FFh							; mode to stall
	jnz		endpoint1_NAK
		mov		A, STALL_IN_OUT
		iowr	ep1_mode
	jmp		endpoint1_done


endpoint1_NAK:
; If the host returned ACK, the SIE set the ACK bit 
; in the epP1_mode register. Clear the ACK bit.
; The endpoint mode remains NAK_IN.
	mov	A, NAK_IN
	iowr	ep1_mode	
	endpoint1_done:
		pop		A
		reti

	
;**********************************************************
;	JUMP TABLES
;**********************************************************

XPAGEOFF

ORG		600h

		; bmRequestTypes commented out are not used for this device,
		; but may be used for your device.  They are kept here as
		; an example of how to use this jumptable.

		bmRequestType_jumptable:
			jmp		h2d_std_device			; 00
			jmp		h2d_std_interface		; 01	
			jmp		h2d_std_endpoint		; 02	
			jmp		request_not_supported	; h2d_std_other			03	
			jmp		request_not_supported	; h2d_class_device		04	
			jmp		h2d_class_interface	;05	
			jmp		request_not_supported	; h2d_class_endpoint	06	
			jmp		request_not_supported	; h2d_class_other		07	
			jmp		request_not_supported	; h2d_vendor_device		08	
			jmp		request_not_supported	; h2d_vendor_interface	09	
			jmp		request_not_supported	; h2d_vendor_endpoint	0A	
			jmp		request_not_supported	; h2d_vendor_other		0B	
			jmp		request_not_supported	; 0C	
			jmp		request_not_supported	; 0D	
			jmp		request_not_supported	; 0E	
			jmp		request_not_supported	; 0F	
			jmp		d2h_std_device			; 10	
			jmp		d2h_std_interface		; 11	
			jmp		d2h_std_endpoint		; 12	
			jmp		request_not_supported	; d2h_std_other			13	
			jmp		request_not_supported	; d2h_class_device		14	
			jmp		d2h_class_interface	; 15	
			jmp		request_not_supported	; d2h_class_endpoint	16	
			jmp		request_not_supported	; d2h_class_other		17	
			jmp		request_not_supported	; d2h_vendor_device		18	
			jmp		request_not_supported	; d2h_vendor_interface	19	
			jmp		request_not_supported	; d2h_vendor_endpoint	1A	
			jmp		request_not_supported	; d2h_vendor_other		1B	
			jmp		request_not_supported	; 1C	
			jmp		request_not_supported	; 1D	
			jmp		request_not_supported	; 1E	
			jmp		request_not_supported	; 1F	

		h2d_std_device_jumptable:
			jmp		request_not_supported	; 00
			jmp		request_not_supported	; 01
			jmp		request_not_supported	; 02
			jmp		request_not_supported		; 03
			jmp		request_not_supported	; 04
			jmp		set_device_address		; 05
			jmp		request_not_supported	; 06

			jmp		request_not_supported	; set_device_descriptor		07
			jmp		request_not_supported	; 08
			jmp		set_device_configuration; 09

		h2d_std_interface_jumptable:
			jmp		request_not_supported	; 00
			jmp		request_not_supported	; clear_interface_feature 01
			jmp		request_not_supported	; 02
			jmp		request_not_supported	; set_interface_feature	 03

			jmp		request_not_supported	; 04
			jmp		request_not_supported	; 05
			jmp		request_not_supported	; 06
			jmp		request_not_supported	; 07
			jmp		request_not_supported	; 08
			jmp		request_not_supported	; 09
			jmp		request_not_supported	; 0A
			jmp		set_interface_interface	; 0B

		h2d_std_endpoint_jumptable:
			jmp		request_not_supported	; 00
			jmp		clear_endpoint_feature	; 01
			jmp		request_not_supported	; 02
			jmp		set_endpoint_feature	; 03

;h2d_class_interface_jumptable was added to logo.asm
;Host-to-device HID-specific requests
		h2d_class_interface_jumptable:
			jmp		request_not_supported	; 00
			jmp		request_not_supported	; 01
			jmp		request_not_supported	; 02
			jmp		request_not_supported	; 03
			jmp		request_not_supported	; 04
			jmp		request_not_supported	; 05
			jmp		request_not_supported	; 06
			jmp		request_not_supported	; 07
			jmp		request_not_supported	; 08

			jmp		set_report			; 09
			jmp		request_not_supported	; 09

			jmp		request_not_supported	; set_idle	; 0A
			jmp		request_not_supported	; set_protocol	; 0B

		d2h_std_device_jumptable:
			jmp		get_device_status		; 00
			jmp		request_not_supported	; 01
			jmp		request_not_supported	; 02
			jmp		request_not_supported	; 03
			jmp		request_not_supported	; 04
			jmp		request_not_supported	; 05
			jmp		get_device_descriptor	; 06
			jmp		request_not_supported	; 07
			jmp		get_device_configuration; 08

		d2h_std_interface_jumptable:
			jmp		get_interface_status	; 00
			jmp		request_not_supported	; 01
			jmp		request_not_supported	; 02
			jmp		request_not_supported	; 03
			jmp		request_not_supported	; 04
			jmp		request_not_supported	; 05
			jmp		get_interface_hid		; 06
			jmp		request_not_supported	; 07
			jmp		request_not_supported	; 08
			jmp		request_not_supported	; 09
			jmp		get_interface_interface	; 0A
	
		d2h_std_endpoint_jumptable:
			jmp		get_endpoint_status		; 00
			jmp		request_not_supported	; 01
			jmp		request_not_supported	; 02
			jmp		request_not_supported	; 03
			jmp		request_not_supported	; 04
			jmp		request_not_supported	; 05
			jmp		request_not_supported	; 06
			jmp		request_not_supported	; 07
			jmp		request_not_supported	; 08
			jmp		request_not_supported	; 09
			jmp		request_not_supported	; 0A
			jmp		request_not_supported	; 0B
			jmp		request_not_supported	; synch frame 0C

;Device-to-host HID-specific requests
		d2h_class_interface_jumptable:
			jmp		request_not_supported	; 00
			jmp		get_report	; 01
			jmp		request_not_supported	; get_idle	; 02
			jmp		request_not_supported	; get_protocol	; 03

		ep0_in_jumptable:
			jmp		request_not_supported
			jmp		control_read_data_stage
			jmp		control_write_status_stage
			jmp		no_data_control_status_stage		

		ep0_out_jumptable:
			jmp		request_not_supported
			jmp		control_read_status_stage
			jmp		control_write_data_stage
			jmp		no_data_control_error

		get_device_descriptor_jumptable:

			jmp		request_not_supported
			jmp		send_device_descriptor
			jmp		send_configuration_descriptor
			jmp		send_string_descriptor
			jmp		send_interface_descriptor
			jmp		send_endpoint_descriptor

		string_jumptable:
			jmp		language_string
			jmp		manufacturer_string
			jmp		product_string

ORG		700h

XPAGEON
		out_report_type_jumptable:
			jmp		request_not_supported ; valid values are 1,2,3
			jmp		request_not_supported ; input report
			jmp		get_output_report ; 2
			jmp		get_feature_report ; 3


		in_report_type_jumptable:
			jmp		request_not_supported ; valid values are 1,2,3
			jmp		send_input_report ; 1
			jmp		request_not_supported  ;output report
			jmp		send_feature_report ; 3


XPAGEON
;*********************************************************
;                   rom lookup tables
;*********************************************************

XPAGEOFF
ORG		800h

control_read_table:
   device_desc_table:
	db	12h			; bLength (18 bytes)
	db	01h			; bDescriptorType (device descriptor)
	db	10h, 01h	; bcdUSB (ver 1.1)
	db	00h			; bDeviceClass (each interface specifies class info)
	db	00h			; bDeviceSubClass (not specified)
	db	00h			; bDeviceProtocol (not specified)
	db	08h			; bMaxPacketSize0 (8 bytes)
	db	77h, 77h	; idVendor (Lakeview Research: 7777h)
	db	34h, 12h	; idProduct (1234h)
	db	00h, 01h	; bcdDevice (1.00) 
	db	01h			; iManufacturer
	db	02h			; iProduct
	db	00h			; iSerialNumber (unused)
	db	01h			; bNumConfigurations (1)

   config_desc_table:
	db	09h			; bLength (9 bytes)
	db	02h			; bDescriptorType (CONFIGURATION)
	db	22h, 00h	; wTotalLength (34 bytes)
	db	01h			; bNumInterfaces (1)
	db	01h			; bConfigurationValue (1)
	db	00h			; Configuration string (unused)
	db	80h			; bmAttributes (bus powered, remote wakeup)
	db	0Dh			; MaxPower (13mA)
   interface_desc_table:
	db	09h			; bLength (9 bytes)
	db	04h			; bDescriptorType (INTERFACE)
	db	00h			; bInterfaceNumber (0)
	db	00h			; bAlternateSetting (0)
	db	01h			; bNumEndpoints (1)
	db	03h			; bInterfaceClass (3..defined by USB spec)
	db	00h			; bInterfaceSubClass (1..defined by USB spec)
	db	00h			; bInterfaceProtocol (2..defined by USB spec)
	db	00h			; iInterface (not supported)
   hid_desc_table:

	db	09h			; bLength (9 bytes)
	db	21h			; bDescriptorType (HID)
	db	00h, 01h	; bcdHID (1.00)	
	db	00h			; bCountryCode (US)
	db	01h			; bNumDescriptors (1)
	db	22h			; bDescriptorType (HID)
	db	end_hid_report_desc_table - hid_report_desc_table, 00h	; wDescriptorLength  
   endpoint_desc_table:
	db	07h			; bLength (7 bytes)
	db	05h			; bDescriptorType (ENDPOINT)
	db	81h			; bEndpointAddress (IN endpoint, endpoint 1)
	db	03h			; bmAttributes (interrupt)
	db	03h, 00h	; wMaxPacketSize (3 bytes)
	db	0Ah			; bInterval (10ms)


   hid_report_desc_table:

	db	06h, A0h, FFh	; usage page (vendor defined)
	db	09h, 01h	; usage (vendor defined)
	db	A1h, 01h	; collection (application)
	db	09h, 02h	; usage (vendor defined)
	db	A1h, 00h	; collection (linked)
	db	06h, A1h, FFh	; usage page (vendor defined)

;The input report
     db 09h, 03h     ;               usage - vendor defined
     db 09h, 04h     ;               usage - vendor defined
     db 15h, 80h     ;               Logical Minimum (-128)
     db 25h, 7Fh     ;               Logical Maximum (127)

     db 35h, 00h     ;               Physical Minimum (0)
     db 45h, FFh;                    Physical Maximum (255)
;    db 66h, 00h, 00h;               Unit (None (2 bytes))
     db 75h, 08h     ;               Report Size (8)  (bits)
     db 95h, 02h     ;               Report Count (2)  (fields)
     db 81h, 02h     ;               Input (Data, Variable, Absolute)  

;The output report
     db 09h, 05h     ;               usage - vendor defined
     db 09h, 06h     ;               usage - vendor defined
     db 15h, 80h     ;               Logical Minimum (-128)
     db 25h, 7Fh     ;               Logical Maximum (127)
     db 35h, 00h     ;               Physical Minimum (0)
     db 45h, FFh     ;               Physical Maximum (255)
;    db 66h, 00h, 00h;               Unit (None (2 bytes))
     db 75h, 08h     ;               Report Size (8)  (bits)
     db 95h, 02h     ;               Report Count (2)  (fields)
     db 91h, 02h     ;               Output (Data, Variable, Absolute)  

;The Feature report
;     db 85h, 01h     ;               report ID
     db 09h, 05h     ;               usage - vendor defined
     db 09h, 06h     ;               usage - vendor defined
     db 15h, 80h     ;               Logical Minimum (-128)
     db 25h, 7Fh     ;               Logical Maximum (127)
     db 35h, 00h     ;               Physical Minimum (0)
     db 45h, FFh     ;               Physical Maximum (255)
;    db 66h, 00h, 00h;               Unit (None (2 bytes))
     db 75h, 08h     ;               Report Size (8)  (bits)
     db 95h, 02h     ;               Report Count (2)  (fields)
     db B1h, 02h     ;               Feature (Data, Variable, Absolute)  


	db	C0h, C0h	; end collection, end collection

    end_hid_report_desc_table:

   device_status_wakeup_disabled:
    db  00h, 00h    ; remote wakeup disabled, bus powered

   device_configured_table:
	db	01h			; device in configured state
   device_unconfigured_table:
	db	00h			; device in unconfigured state


   endpoint_nostall_table:
	db	00h, 00h	; endpoint not stalled
   endpoint_stall_table:
	db	01h, 00h	; endpoint stalled

   interface_status_table:
	db	00h, 00h	; default response


   interface_alternate_table:
	db	00h			; only valid alternate setting

   interface_boot_protocol:
	db	00h
   interface_report_protocol:
	db	01h

   ilanguage_string:
    db 04h          				; Length
    db 03h          				; Type (3=string)

    db 09h          				; Language:  English
    db 04h          				; Sub-language: US

   imanufacturer_string:
;Length is (# of characters + 1) * 2
    db 38h          				; Length
    db 03h          				; Type (3=string)
    dsu "USB HID Design by USBLab"

   iproduct_string:
    db 18h          				; Length
    db 03h          				; Type (3=string)
    dsu "HID Example"

XPAGEOFF


;**********************************************************
;	APPLICATION SPECIFIC TABLES
;**********************************************************
XPAGEON

ORG		900h


		event_machine_jumptable:
			jmp		no_event_pending
			jmp		no_event_task

XPAGEON
