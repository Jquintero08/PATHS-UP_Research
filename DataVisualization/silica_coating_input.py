import tkinter as tk
import tkinter.font
#import configparser
import time
import board
from adafruit_motorkit import MotorKit
from gpiozero import Buzzer
#from gpiozero.pins.pigpio import PiGPIOFactory
import threading

buzzer = Buzzer(23)
win=tk.Tk()
win.title("Nanostar")
win.geometry('570x240')
HeaderFont=tkinter.font.Font(family = 'Helvetica', size = 18, weight = "bold")
myFont=tkinter.font.Font(family = 'Helvetica', size = 12, weight = "bold")
myFont1=tkinter.font.Font(family = 'Helvetica', size = 10, weight = "bold")
kit = MotorKit(i2c=board.I2C())
kit2 = MotorKit(address = 0x61)


def PrimeAll():
    kit.motor1.throttle = 1
    kit.motor2.throttle = 1
    kit.motor3.throttle = 1
    kit.motor4.throttle = 1
    kit2.motor1.throttle = 1
    kit2.motor2.throttle = 1
    kit2.motor3.throttle = 1
    kit2.motor4.throttle = 1
    time.sleep(3)
    kit.motor1.throttle = 0
    kit.motor2.throttle = 0
    kit.motor3.throttle = 0
    kit.motor4.throttle = 0
    kit2.motor1.throttle = 0
    kit2.motor2.throttle = 0
    kit2.motor3.throttle = 0
    kit2.motor4.throttle = 0
    
def Prime1():
    #time = tkinter.simpledialog.askinteger("Input", "Enter how much (mL)") # Will just get the time in seconds
    #amount = time*1.3 # Will turn seconds into mL (6.5 mL = 5s * 1.3)
    kit.motor1.throttle = 1
    time.sleep(3) # Change to 3 if you don't want input or amount and uncomment if you do
    kit.motor1.throttle = 0
    
def Prime2():
    kit.motor2.throttle = 1
    time.sleep(3)
    kit.motor2.throttle = 0
    
def Prime3():
    kit.motor3.throttle = 1
    time.sleep(3)
    kit.motor3.throttle = 0
    
def Prime4():
    kit.motor4.throttle = 1
    time.sleep(3)
    kit.motor4.throttle = 0
    
def Prime5():
    kit2.motor1.throttle = 1
    time.sleep(3)
    kit2.motor1.throttle = 0
    
def Prime6():
    kit2.motor2.throttle = 1
    time.sleep(3)
    kit2.motor2.throttle = 0
    
def Prime7():
    kit2.motor3.throttle = 1
    time.sleep(3)
    kit2.motor3.throttle = 0
    
def Prime8():
    kit2.motor4.throttle = 1
    time.sleep(3)
    kit2.motor4.throttle = 0

def run_motor2():
    kit.motor2.throttle = 1 #add AgNO3 (PUMP 2)
    time.sleep(0.96)
    kit.motor2.throttle = 0

def run_motor3():
    kit.motor3.throttle = 1 #ad AA (PUMP 3)
    time.sleep(0.89)
    kit.motor3.throttle = 0

def Nanostar():
    time.sleep(2)
    kit.motor1.throttle = 1 #add seed (PUMP 1)
    time.sleep(1.23)
    kit.motor1.throttle = 0
    time.sleep(30)
    
    t2 = threading.Thread(target=kit.motor2)
    t2.start()
    
    t3 = threading.Thread(target=kit.motor3)
    t3.start()
    
    t2.join()
    t3.join()
    
    time.sleep(120)
    
    buzzer.on()
    time.sleep(1)
    buzzer.off()
    time.sleep(0.5)
    buzzer.on()
    time.sleep(1)
    buzzer.off()
    time.sleep(0.5)
    buzzer.on()
    time.sleep(1)
    buzzer.off()
    
    
    
def Silica_coating():
    time.sleep(2)
    kit2.motor1.throttle = 1 #add 12 ml of Solution A (PUMP 5)
    time.sleep(12)
    kit2.motor1.throttle = 0
    kit2.motor2.throttle = 1 #add 4 ml of EtOH (PUMP 6)
    time.sleep(4)
    kit2.motor2.throttle = 0
    kit2.motor3.throttle = 1 #add 400 uL of star (PUMP 7)
    time.sleep(0.4)
    kit2.motor4.throttle = 0
    kit2.motor4.throttle = 1 #add 200 uL of NH3 in H2O (PUMP 8)
    time.sleep(0.2)
    kit2.motor4.throttle = 0
    time.sleep(900)
    buzzer.on()
    time.sleep(1)
    buzzer.off()
    time.sleep(0.5)
    buzzer.on()
    time.sleep(1)
    buzzer.off()
    time.sleep(0.5)
    buzzer.on()
    time.sleep(1)
    buzzer.off()

def exitProgram():
    win.destroy()
    
PrimeButton=tk.Button(win, text='Prime All Pumps', font=myFont, command=PrimeAll, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=1, row=0, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Prime Pump 1', font=myFont, command=Prime1, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=0, row=1, sticky=tk.NSEW)

#PrimeButton=tk.Button(win, text='Prime Pump 2', font=myFont, command=Prime2, bg='bisque2', height=1, width=18)
#PrimeButton.grid(column=1, row=1, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Prime Pump 3', font=myFont, command=Prime3, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=2, row=1, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Prime Pump 4', font=myFont, command=Prime4, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=0, row=2, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Prime Pump 5', font=myFont, command=Prime5, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=1, row=2, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Prime Pump 6', font=myFont, command=Prime6, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=2, row=2, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Prime Pump 7', font=myFont, command=Prime7, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=0, row=3, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Prime Pump 8', font=myFont, command=Prime8, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=1, row=3, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Nanostar', font=myFont, command=Nanostar, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=2, row=3, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Silica Coating', font=myFont, command=Silica_coating, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=0, row=4, sticky=tk.NSEW)

PrimeButton=tk.Button(win, text='Exit', font=myFont, command=exitProgram, bg='bisque2', height=1, width=18)
PrimeButton.grid(column=1, row=6, sticky=tk.NSEW)