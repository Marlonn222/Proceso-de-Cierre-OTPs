import time
import pyautogui
from time import sleep
import pyperclip
from storagefunctions import (pressingKey,selectToEnd,formatDate)

def searchCloseOTPs(incidentId,usuario,cod_resolucion_1,BillingItem,anotaciones):
    
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = anotaciones_ot_field = crm_warning_message = crm_ot_blocked_message = mod_consulta_popup = None 
    OTP_Estado1 = OTP_Estado2 = estado_field  = cod_resolucion_1_field = None
    crmAttempts = 0
    print("BillingItem: ",BillingItem)
    print("incidentId: ",incidentId)

    #/////////////////////////////////// CRM DASHBOARD ///////////////////////////////////////////    
    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/Automatizacion/assets/crm_dashboard.png', grayscale = True,confidence=0.85)
        sleep(0.5)
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].minimize()
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].maximize()
            print("Estoy dentro de los 5 intentos para ver el Dashboard del CRM")
            crmAttempts = 0
        crmAttempts = crmAttempts + 1
                
    print("CRM Dashboard GUI is present!")   
    pyautogui.click(pyautogui.center(crm_dashboard))    
    
    sleep(0.5)
    pressingKey('f2')
    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/Automatizacion/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(0.5)
    pressingKey('enter')
    
    #/////////////////////////////////// FASE DE ADVERTENCIAS ///////////////////////////////////////////
    
    while crm_ot_blocked_message is None and crm_warning_message is None and crmAttempts < 10:
        crm_warning_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_ot_blocked_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/crm_ot_blocked_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up de OT bloqueada inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)            
            
    if(crm_ot_blocked_message is None and crm_warning_message is None):
        print("No se encuentran ventanas de advertencia!")                           
        
        # Validate edit_incident view is visible and on focus            
        while crm_edit_incident is None:
            crm_edit_incident = pyautogui.locateOnScreen('C:/Automatizacion/assets/edit_incident.png', grayscale = True,confidence=0.9)   
        print("Vista de Detalles is present!")    
        print("CRM Edit Incident button is present!")
        crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
        pyautogui.click(crm_edit_incident_x, crm_edit_incident_y)
    
        # Validate assign user pop up is visible and on focus
        # crm_assign_user = None  # reset variable   
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)
    
    else:
        pressingKey('enter')
        sleep(3)
        pressingKey('enter')
        sleep(1)
        pyautogui.hotkey('alt','f4')
        return 9
    
    # Maximize CRM window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    print("CRM Edit view Window was maximized!")    
    # Click on open Details Button on the CRM to make date fields visible    
    pyautogui.click(19,493)
    crmAttempts = 0
    
    #/////////////////////////////////// TRASLADO DE chek list /////////////////////////////////////    
    while anotaciones_ot_field is None:
        print("buscando anotaciones_ot_field in screen")
        anotaciones_ot_field = pyautogui.locateOnScreen('C:/Automatizacion/assets/anotaciones_ot_field.png', grayscale = True,confidence=0.95)                                
   
    print("anotaciones_ot_field is present!")
    pyautogui.moveTo(pyautogui.center(anotaciones_ot_field))            
    pyautogui.click()
    pressingKey('tab')
    sleep(0.5)
    pressingKey('tab')
    pyperclip.copy(anotaciones)
    sleep(0.5)
    pyautogui.hotkey('ctrl','v')    
    sleep(1)    

    #/////////////////////////////////// ITEM DE FACTURACION  /////////////////////////////////////  
        
    pyautogui.click(303,497)
    pyautogui.press('tab',presses=42)   
    sleep(1)
    pyautogui.write("NAE")
    pressingKey('tab')
    pyautogui.write("NAE")
    pressingKey('tab')
    pyautogui.write("GESTION")    
    pressingKey('tab')
    pyautogui.write(BillingItem)
    pressingKey('tab')
    sleep(1)
    
    # CAMBIO DE CODIGO DE RESOLUCION 1
    while cod_resolucion_1_field is None:
        print("buscando cod_resolucion_1_field in screen")
        cod_resolucion_1_field = pyautogui.locateOnScreen('C:/Automatizacion/assets/cod_resolucion_1_field.png', grayscale = True,confidence=0.95)        
    print("cod_resolucion_1_field is present!")
    
    cod_resolucion_1_field_x,cod_resolucion_1_field_y = pyautogui.center(cod_resolucion_1_field)    
    pyautogui.moveTo(cod_resolucion_1_field_x,cod_resolucion_1_field_y)
    pyautogui.moveRel(181,0)     
    pyautogui.doubleClick()
    sleep(1)    
    pyautogui.write(cod_resolucion_1)
    pressingKey('enter')

    #/////////////////////////////////// CAMBIO DE ESTADO  /////////////////////////////////////
    
    # Bucle para buscar otros estados
    while OTP_Estado1 is None or OTP_Estado2 is None:
        print("Buscando Estados en la pantalla...")
        OTP_Estado1 = pyautogui.locateOnScreen('C:/Automatizacion/assets/Estado_Ejecucion.png', grayscale=True, confidence=0.95)
        OTP_Estado2 = pyautogui.locateOnScreen('C:/Automatizacion/assets/Estado_Terminada.png', grayscale=True, confidence=0.95)
        if OTP_Estado1 is not None:
            #/////////////// CABIO DE ESTADO DE EJECUCION A TERMINADA ///////////////////////////////////
            while estado_field is None:
                print("buscando estado_field in screen")
                estado_field = pyautogui.locateOnScreen('C:/Automatizacion/assets/estado_field.png', grayscale = True,confidence=0.95)        
            print("estado_field is present!")
            pyautogui.moveTo(pyautogui.center(estado_field))
            pyautogui.click()
            sleep(0.5)
            posicion_actual = pyautogui.position()
            print("Posición actual del cursor:", posicion_actual)
            # Mover el cursor 176 píxeles a la derecha desde su posición actual
            pyautogui.moveRel(176, 0, duration=0.5)
            # Imprimir la nueva posición del cursor
            nueva_posicion = pyautogui.position()
            print("Nueva posición del cursor:", nueva_posicion)
            pyautogui.doubleClick(nueva_posicion)
            sleep(1)
            pyautogui.press('up')
            sleep(0.5)
            pyautogui.press('tab')
            pyautogui.write(usuario)
            pyautogui.press('enter')
            sleep(1)
            #///////////////////////////////////////////////////////////////////////////////////////////
            # Proceso de guardado dando clicl en el boton
            while crm_save_incident is None:
                print("buscando crm_save_incident in screen")
                crm_save_incident = pyautogui.locateOnScreen('C:/Automatizacion/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
            print("CRM Save Incident button is present!")
            crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
            sleep(0.5)
            pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
            sleep(1)
            crmAttempts = 0
            while crm_warning_message is None and crmAttempts < 10:        
                crm_warning_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/crm_process_message.png', grayscale = True,confidence=0.9)   
                sleep(0.5)
                if crmAttempts == 8:
                    print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up inesperado en el CRM")                        
                crmAttempts += 1
                print(crmAttempts)
        
            if(crm_warning_message is None):
                print("CRM empty_warning_message_crm pop up is not present!")                    
        
            # Validate assign user pop up is visible and on focus   
            crm_assign_user = None  # reset variable
            crmAttempts = 0
            while crm_assign_user is None and crmAttempts < 5:
                crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
                sleep(0.5)
                print("buscando ventana de asignación de usuario REVISADO")
                print("CRM ASSIGN USER:",crm_assign_user)
                crmAttempts +=1
            if crm_assign_user is not None:
                print("CRM Confirm Assign Pop Up is present!")
                sleep(1)
                pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
                sleep(1)

            # Validate save OTP successfully  pop up is visible and on focus   
            while crm_otp_saved_sucessfully is None:
                print("buscando crm_otp_saved_sucessfully en pantalla")
                crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/Automatizacion/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
            print("CRM OTP Saved Succesfully Pop Up is present!")
            sleep(1)
            pressingKey('enter')
            sleep(1)
            #/////////////// CABIO DE ESTADO DE TERMINADA A CERRADA ///////////////////////////////////
            while estado_field is None:
                print("buscando estado_field in screen")
                estado_field = pyautogui.locateOnScreen('C:/Automatizacion/assets/estado_field.png', grayscale = True,confidence=0.95)        
            print("estado_field is present!")
            pyautogui.moveTo(pyautogui.center(estado_field))
            pyautogui.click()
            sleep(0.5)
            posicion_actual = pyautogui.position()
            print("Posición actual del cursor:", posicion_actual)
            # Mover el cursor 176 píxeles a la derecha desde su posición actual
            pyautogui.moveRel(176, 0, duration=0.5)
            # Imprimir la nueva posición del cursor
            nueva_posicion = pyautogui.position()
            print("Nueva posición del cursor:", nueva_posicion)
            pyautogui.doubleClick(nueva_posicion)
            sleep(1)
            pyautogui.press('up')
            sleep(0.5)
            pyautogui.press('tab')
            pyautogui.write(usuario)
            pyautogui.press('enter')
            sleep(1)
            break  # Salir del bucle una vez que se encuentra y guarda la variable volatil
        if OTP_Estado2 is not None:
            #/////////////// CABIO DE ESTADO DE TERMINADA A CERRADA ///////////////////////////////////
            while estado_field is None:
                print("buscando estado_field in screen")
                estado_field = pyautogui.locateOnScreen('C:/Automatizacion/assets/estado_field.png', grayscale = True,confidence=0.95)        
            print("estado_field is present!")
            pyautogui.moveTo(pyautogui.center(estado_field))
            pyautogui.click()
            sleep(0.5)
            posicion_actual = pyautogui.position()
            print("Posición actual del cursor:", posicion_actual)
            # Mover el cursor 176 píxeles a la derecha desde su posición actual
            pyautogui.moveRel(176, 0, duration=0.5)
            # Imprimir la nueva posición del cursor
            nueva_posicion = pyautogui.position()
            print("Nueva posición del cursor:", nueva_posicion)
            pyautogui.doubleClick(nueva_posicion)
            sleep(1)
            pyautogui.press('up')
            sleep(0.5)
            pyautogui.press('tab')
            pyautogui.write(usuario)
            pyautogui.press('enter')
            sleep(1)
            break  # Salir del bucle una vez que se encuentra y guarda la variable volatil
    # Realizar un proceso si se encuentra cualquiera de los otros estados
    print("Estado encontrado")
  
    # Proceso de guardado dando clicl en el boton y culminar proceso
    while crm_save_incident is None:
        print("buscando crm_save_incident in screen")
        crm_save_incident = pyautogui.locateOnScreen('C:/Automatizacion/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    sleep(0.5)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
    sleep(1)
    crmAttempts = 0
    while crm_warning_message is None and crmAttempts < 10:        
        crm_warning_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/crm_process_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)
        
    if(crm_warning_message is None):
        print("CRM empty_warning_message_crm pop up is not present!")                    
        
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario REVISADO")
            print("CRM ASSIGN USER:",crm_assign_user)
            crmAttempts +=1
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

        # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None:
            print("buscando crm_otp_saved_sucessfully en pantalla")
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/Automatizacion/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
        print("CRM OTP Saved Succesfully Pop Up is present!")
        sleep(1)
        pressingKey('enter')
        sleep(1)

    #/////////////////////////////////// CLOSING PHASE /////////////////////////////////////    
    
    while crm_save_incident is None:
        crm_save_incident = pyautogui.locateOnScreen('C:/Automatizacion/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
    sleep(1)    
    crm_assign_user = None
    while crm_warning_message is None and crm_assign_user is None:  
        print("buscando o crm_warning_message o crm_assign_user")      
        crm_warning_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
        sleep(0.5)                
        
    if(crm_warning_message is None):
        print("CRM empty_warning_message_crm pop up is not present!")
                
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")            
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

        # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None:
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/Automatizacion/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
        print("CRM OTP Saved Succesfully Pop Up is present!")
        sleep(1)
        pressingKey('enter')
        sleep(1)

        # Validate if changes must be saved or not and Close the OT Details View On Edition mode  
        # Close Edit Incident View 
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        return 0
    else:
        print("CRM empty_warning_message_crm pop up is present!")
        #pyautogui.click(pyautogui.center(crm_warning_message))
        sleep(0.5)
        pressingKey('enter')
        sleep(1)
        pressingKey('enter')
        sleep(1)
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1)
        pressingKey('n')        
            
        return 9
        
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        
        # You must Validate assign user pop up is visible and on focus
        sleep(1)
        pressingKey('n') # No re asignar usuario al momento de editar las OTPs         
        print("n key has been pressed!")         
def searchCloseOTPsRUN(incidentId,usuario, anotaciones):
    
    crm_dashboard = crm_assign_user = crm_edit_incident = crm_save_incident = crm_otp_saved_sucessfully = anotaciones_ot_field = crm_warning_message = crm_ot_blocked_message = mod_consulta_popup = None 
    OTP_Estado1 = OTP_Estado2 = estado_field  = cod_resolucion_1_field = None
    crmAttempts = 0
    print("incidentId: ",incidentId)

    #/////////////////////////////////// CRM DASHBOARD ///////////////////////////////////////////    
    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/Automatizacion/assets/crm_dashboard.png', grayscale = True,confidence=0.85)
        sleep(0.5)
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].minimize()
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].maximize()
            print("Estoy dentro de los 5 intentos para ver el Dashboard del CRM")
            crmAttempts = 0
        crmAttempts = crmAttempts + 1
                
    print("CRM Dashboard GUI is present!")   
    pyautogui.click(pyautogui.center(crm_dashboard))    
    
    sleep(0.5)
    pressingKey('f2')
    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/Automatizacion/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(0.5)
    pressingKey('enter')
    
    #/////////////////////////////////// FASE DE ADVERTENCIAS ///////////////////////////////////////////
    
    while crm_ot_blocked_message is None and crm_warning_message is None and crmAttempts < 10:
        crm_warning_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_ot_blocked_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/crm_ot_blocked_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up de OT bloqueada inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)            
            
    if(crm_ot_blocked_message is None and crm_warning_message is None):
        print("No se encuentran ventanas de advertencia!")                           
        
        # Validate edit_incident view is visible and on focus            
        while crm_edit_incident is None:
            crm_edit_incident = pyautogui.locateOnScreen('C:/Automatizacion/assets/edit_incident.png', grayscale = True,confidence=0.9)   
        print("Vista de Detalles is present!")    
        print("CRM Edit Incident button is present!")
        crm_edit_incident_x,crm_edit_incident_y = pyautogui.center(crm_edit_incident)
        pyautogui.click(crm_edit_incident_x, crm_edit_incident_y)
    
        # Validate assign user pop up is visible and on focus
        # crm_assign_user = None  # reset variable   
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)
    
    else:
        pressingKey('enter')
        sleep(3)
        pressingKey('enter')
        sleep(1)
        pyautogui.hotkey('alt','f4')
        return 9
    
    # Maximize CRM window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    print("CRM Edit view Window was maximized!")    
    # Click on open Details Button on the CRM to make date fields visible    
    pyautogui.click(19,493)
    crmAttempts = 0
    
    #/////////////////////////////////// TRASLADO DE chek list /////////////////////////////////////    
    while anotaciones_ot_field is None:
        print("buscando anotaciones_ot_field in screen")
        anotaciones_ot_field = pyautogui.locateOnScreen('C:/Automatizacion/assets/anotaciones_ot_field.png', grayscale = True,confidence=0.95)                                
   
    print("anotaciones_ot_field is present!")
    pyautogui.moveTo(pyautogui.center(anotaciones_ot_field))            
    pyautogui.click()
    pressingKey('tab')
    sleep(0.5)
    pressingKey('tab')
    pyperclip.copy(anotaciones)
    sleep(0.5)
    pyautogui.hotkey('ctrl','v')    
    sleep(1)    

    #/////////////////////////////////// CAMBIO DE ESTADO  /////////////////////////////////////
    
    # Bucle para buscar otros estados
    while OTP_Estado1 is None or OTP_Estado2 is None:
        print("Buscando Estados en la pantalla...")
        OTP_Estado1 = pyautogui.locateOnScreen('C:/Automatizacion/assets/Estado_Ejecucion.png', grayscale=True, confidence=0.95)
        OTP_Estado2 = pyautogui.locateOnScreen('C:/Automatizacion/assets/Estado_Terminada.png', grayscale=True, confidence=0.95)
        if OTP_Estado1 is not None:
            #/////////////// CABIO DE ESTADO DE EJECUCION A TERMINADA ///////////////////////////////////
            while estado_field is None:
                print("buscando estado_field in screen")
                estado_field = pyautogui.locateOnScreen('C:/Automatizacion/assets/estado_field.png', grayscale = True,confidence=0.95)        
            print("estado_field is present!")
            pyautogui.moveTo(pyautogui.center(estado_field))
            pyautogui.click()
            sleep(0.5)
            posicion_actual = pyautogui.position()
            print("Posición actual del cursor:", posicion_actual)
            # Mover el cursor 176 píxeles a la derecha desde su posición actual
            pyautogui.moveRel(176, 0, duration=0.5)
            # Imprimir la nueva posición del cursor
            nueva_posicion = pyautogui.position()
            print("Nueva posición del cursor:", nueva_posicion)
            pyautogui.doubleClick(nueva_posicion)
            sleep(1)
            pyautogui.press('up')
            sleep(0.5)
            pyautogui.press('tab')
            pyautogui.write(usuario)
            pyautogui.press('enter')
            sleep(1)
            #///////////////////////////////////////////////////////////////////////////////////////////
            # Proceso de guardado dando clicl en el boton
            while crm_save_incident is None:
                print("buscando crm_save_incident in screen")
                crm_save_incident = pyautogui.locateOnScreen('C:/Automatizacion/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
            print("CRM Save Incident button is present!")
            crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
            sleep(0.5)
            pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
            sleep(1)
            crmAttempts = 0
            while crm_warning_message is None and crmAttempts < 10:        
                crm_warning_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/crm_process_message.png', grayscale = True,confidence=0.9)   
                sleep(0.5)
                if crmAttempts == 8:
                    print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up inesperado en el CRM")                        
                crmAttempts += 1
                print(crmAttempts)
        
            if(crm_warning_message is None):
                print("CRM empty_warning_message_crm pop up is not present!")                    
        
            # Validate assign user pop up is visible and on focus   
            crm_assign_user = None  # reset variable
            crmAttempts = 0
            while crm_assign_user is None and crmAttempts < 5:
                crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
                sleep(0.5)
                print("buscando ventana de asignación de usuario REVISADO")
                print("CRM ASSIGN USER:",crm_assign_user)
                crmAttempts +=1
            if crm_assign_user is not None:
                print("CRM Confirm Assign Pop Up is present!")
                sleep(1)
                pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
                sleep(1)

            # Validate save OTP successfully  pop up is visible and on focus   
            while crm_otp_saved_sucessfully is None:
                print("buscando crm_otp_saved_sucessfully en pantalla")
                crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/Automatizacion/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
            print("CRM OTP Saved Succesfully Pop Up is present!")
            sleep(1)
            pressingKey('enter')
            sleep(1)
            #/////////////// CABIO DE ESTADO DE TERMINADA A CERRADA ///////////////////////////////////
            while estado_field is None:
                print("buscando estado_field in screen")
                estado_field = pyautogui.locateOnScreen('C:/Automatizacion/assets/estado_field.png', grayscale = True,confidence=0.95)        
            print("estado_field is present!")
            pyautogui.moveTo(pyautogui.center(estado_field))
            pyautogui.click()
            sleep(0.5)
            posicion_actual = pyautogui.position()
            print("Posición actual del cursor:", posicion_actual)
            # Mover el cursor 176 píxeles a la derecha desde su posición actual
            pyautogui.moveRel(176, 0, duration=0.5)
            # Imprimir la nueva posición del cursor
            nueva_posicion = pyautogui.position()
            print("Nueva posición del cursor:", nueva_posicion)
            pyautogui.doubleClick(nueva_posicion)
            sleep(1)
            pyautogui.press('up')
            sleep(0.5)
            pyautogui.press('tab')
            pyautogui.write(usuario)
            pyautogui.press('enter')
            sleep(1)
            break  # Salir del bucle una vez que se encuentra y guarda la variable volatil
        if OTP_Estado2 is not None:
            #/////////////// CABIO DE ESTADO DE TERMINADA A CERRADA ///////////////////////////////////
            while estado_field is None:
                print("buscando estado_field in screen")
                estado_field = pyautogui.locateOnScreen('C:/Automatizacion/assets/estado_field.png', grayscale = True,confidence=0.95)        
            print("estado_field is present!")
            pyautogui.moveTo(pyautogui.center(estado_field))
            pyautogui.click()
            sleep(0.5)
            posicion_actual = pyautogui.position()
            print("Posición actual del cursor:", posicion_actual)
            # Mover el cursor 176 píxeles a la derecha desde su posición actual
            pyautogui.moveRel(176, 0, duration=0.5)
            # Imprimir la nueva posición del cursor
            nueva_posicion = pyautogui.position()
            print("Nueva posición del cursor:", nueva_posicion)
            pyautogui.doubleClick(nueva_posicion)
            sleep(1)
            pyautogui.press('up')
            sleep(0.5)
            pyautogui.press('tab')
            pyautogui.write(usuario)
            pyautogui.press('enter')
            sleep(1)
            break  # Salir del bucle una vez que se encuentra y guarda la variable volatil
    # Realizar un proceso si se encuentra cualquiera de los otros estados
    print("Estado encontrado")
  
    # Proceso de guardado dando clicl en el boton y culminar proceso
    while crm_save_incident is None:
        print("buscando crm_save_incident in screen")
        crm_save_incident = pyautogui.locateOnScreen('C:/Automatizacion/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    sleep(0.5)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
    sleep(1)
    crmAttempts = 0
    while crm_warning_message is None and crmAttempts < 10:        
        crm_warning_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/crm_process_message.png', grayscale = True,confidence=0.9)   
        sleep(0.5)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)
        
    if(crm_warning_message is None):
        print("CRM empty_warning_message_crm pop up is not present!")                    
        
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario REVISADO")
            print("CRM ASSIGN USER:",crm_assign_user)
            crmAttempts +=1
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

        # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None:
            print("buscando crm_otp_saved_sucessfully en pantalla")
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/Automatizacion/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
        print("CRM OTP Saved Succesfully Pop Up is present!")
        sleep(1)
        pressingKey('enter')
        sleep(1)

    #/////////////////////////////////// CLOSING PHASE /////////////////////////////////////    
    
    while crm_save_incident is None:
        crm_save_incident = pyautogui.locateOnScreen('C:/Automatizacion/assets/guardar_incidente_button.png', grayscale = True,confidence=0.9)   
    print("CRM Save Incident button is present!")
    crm_save_incident_x,crm_save_incident_y = pyautogui.center(crm_save_incident)
    pyautogui.click(crm_save_incident_x, crm_save_incident_y)   
    
    sleep(1)    
    crm_assign_user = None
    while crm_warning_message is None and crm_assign_user is None:  
        print("buscando o crm_warning_message o crm_assign_user")      
        crm_warning_message = pyautogui.locateOnScreen('C:/Automatizacion/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
        sleep(0.5)                
        
    if(crm_warning_message is None):
        print("CRM empty_warning_message_crm pop up is not present!")
                
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")            
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

        # Validate save OTP successfully  pop up is visible and on focus   
        while crm_otp_saved_sucessfully is None:
            crm_otp_saved_sucessfully = pyautogui.locateOnScreen('C:/Automatizacion/assets/incidente_guardado_exitosamente.png', grayscale = True,confidence=0.9)   
        print("CRM OTP Saved Succesfully Pop Up is present!")
        sleep(1)
        pressingKey('enter')
        sleep(1)

        # Validate if changes must be saved or not and Close the OT Details View On Edition mode  
        # Close Edit Incident View 
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        return 0
    else:
        print("CRM empty_warning_message_crm pop up is present!")
        #pyautogui.click(pyautogui.center(crm_warning_message))
        sleep(0.5)
        pressingKey('enter')
        sleep(1)
        pressingKey('enter')
        sleep(1)
        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1)
        pressingKey('n')        
            
        return 9
        
        # Validate assign user pop up is visible and on focus   
        crm_assign_user = None  # reset variable
        crmAttempts = 0
        while crm_assign_user is None and crmAttempts < 5:
            crm_assign_user = pyautogui.locateOnScreen('C:/Automatizacion/assets/asignar_ot_usuario_operador.png', grayscale = True,confidence=0.9)   
            sleep(0.5)
            print("buscando ventana de asignación de usuario")
            crmAttempts +=1            
        if crm_assign_user is not None:
            print("CRM Confirm Assign Pop Up is present!")
            sleep(1)
            pressingKey('n') # No re asignar usuario al momento de editar las OTPs 
            sleep(1)

        pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
        sleep(1.5)
        
        # You must Validate assign user pop up is visible and on focus
        sleep(1)
        pressingKey('n') # No re asignar usuario al momento de editar las OTPs         
        print("n key has been pressed!")         
