import serial
import struct
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os
import threading
import queue
import time
from tkinter import Tk, Button, Label, messagebox, Frame
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re

class AccelerometerGUI:
    def __init__(self, master):
        # Iniciar comunicación serie con Arduino
        self.data = []
        self.master = master
        self.master.title("Accelerometer Data Collector")

        # Marco para los botones
        self.button_frame = Frame(master)
        self.button_frame.pack(side="top", padx=10, pady=8)

        self.s = None
        self.running = False
        self.data_queue = queue.Queue()

        # Figura con subgráficas: 3 filas, 2 columnas
        self.fig, self.axs = plt.subplots(3, 2, figsize=(18, 12))
        self.lines = []
        self.fft_lines = []
        self.colors = ['r', 'g', 'b']
        self.labels = ['X', 'Y', 'Z']

        plt.subplots_adjust(hspace=0.4, wspace=0.3)

        for i in range(3):
            line, = self.axs[i, 0].plot([], [], color=self.colors[i], label=f'Acceleration {self.labels[i]}')
            self.axs[i, 0].set_ylabel('g')
            self.axs[i, 0].set_title(f'Acceleration {self.labels[i]}')
            self.axs[i, 0].legend(loc='upper right')
            self.lines.append(line)

            fft_line, = self.axs[i, 1].plot([], [], color=self.colors[i], label=f'FFT {self.labels[i]}')
            self.axs[i, 1].set_ylabel('Log Magnitude')
            self.axs[i, 1].set_title(f'FFT {self.labels[i]}')
            self.axs[i, 1].legend(loc='upper right')
            self.fft_lines.append(fft_line)

        self.axs[2, 0].set_xlabel('Time (s)')
        self.axs[2, 1].set_xlabel('Frequency (Hz)')

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.master)
        self.canvas.get_tk_widget().pack(side="right", padx=10, pady=10)

        # Botones
        self.start_button = Button(self.button_frame, text="Iniciar toma de datos", command=self.start_data_collection, width=20, height=1)
        self.start_button.pack(side="left", padx=10, pady=8)

        self.stop_button = Button(self.button_frame, text="Detener toma de datos", command=self.stop_data_collection, width=20, height=1)
        self.stop_button.pack(side="left", padx=10)

        self.save_button = Button(self.button_frame, text="Guardar datos en Excel", command=self.save_data, state="disabled", width=20, height=1)
        self.save_button.pack(side="left", padx=10)

        self.exit_button = Button(self.button_frame, text="Salir", command=self.exit_application, width=10, height=1)
        self.exit_button.pack(side="left", padx=10)

        self.status_label = Label(self.button_frame, text="", width=25, height=1)
        self.status_label.pack(side="left", padx=10)
        
        self.corriente = Label(self.button_frame, text="", width=20, height=1)
        self.corriente.pack(side="left", padx=10)

    def start_data_collection(self):
        if not self.running:
            self.aa = np.zeros((1, 3))
            self.tt = np.array([0])
            self.start_time = None

            try:
                self.arduino = serial.Serial('COM5', 9600)  # Ajusta 'COM3' según tu configuración
                self.ser = serial.Serial('COM7', 115200)  # Ajusta 'COM7' según tu configuración
                time.sleep(2)  # Espera a que las conexiones se establezcan
                self.running = True
                self.start_time = time.time()
                self.data_thread = threading.Thread(target=self.collect_data)
                self.data_thread.start()
                self.update_graph()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to start data collection: {e}")

    def stop_data_collection(self):
        if self.running:
            self.running = False
            if self.data_thread.is_alive():
                self.data_thread.join()
            if self.ser and self.ser.is_open:
                self.ser.close()

    def save_data(self):
        try:
            if len(self.tt) > 15:
                data_to_save = self.aa[15:]
                time_to_save = self.tt[15:]
                aa_to_plot = self.aa[15:]

                # Convertir los datos de aa en un DataFrame de pandas
                df_accel = pd.DataFrame(data_to_save, columns=['X', 'Y', 'Z'])
                df_accel['Tiempo (s)'] = time_to_save  # Añadir columna de tiempo
                
                # Calcular FFT
                aa_to_plot_centered = aa_to_plot - np.mean(aa_to_plot, axis=0)
                fft_x = np.fft.fft(aa_to_plot_centered[:, 0])
                fft_y = np.fft.fft(aa_to_plot_centered[:, 1])
                fft_z = np.fft.fft(aa_to_plot_centered[:, 2])
                
                # Obtener las frecuencias correspondientes
                freqs = np.fft.fftfreq(len(data_to_save), d=0.01)  # Frecuencias (suponiendo que la frecuencia de muestreo es 100Hz)

                # Agregar los resultados de la FFT al DataFrame
                
                df_accel['FFT_X'] = np.abs(fft_x)
                df_accel['FFT_Y'] = np.abs(fft_y)
                df_accel['FFT_Z'] = np.abs(fft_z)

                # Crear un DataFrame para los datos del sensor de corriente
                df_current = pd.DataFrame({'Corriente': self.data})

                # Combinar los DataFrames en uno solo
                df_combined = pd.concat([df_accel.reset_index(drop=True), df_current.reset_index(drop=True)], axis=1)

                # Crear un nombre de archivo con fecha y hora
                timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")  # Formato de fecha y hora
                filename = f'acelerometro_datos_{timestamp}.xlsx'  # Nombre de archivo con la fecha y hora
                df_combined.to_excel(filename, index=False)  # Guardar datos en Excel

                # Obtener y mostrar la ubicación completa del archivo
                file_path = os.path.join(os.getcwd(), filename)
                messagebox.showinfo("Success", f'Data guardada en: {file_path}')

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {e}")

    def exit_application(self):
        self.stop_data_collection()
        plt.close(self.fig)
        self.master.quit()
        self.master.destroy()

    def collect_data(self):
        try:
            self.contador = 0
            while self.running:
                if self.arduino.in_waiting > 0:
                    line = self.arduino.readline().decode('utf-8').rstrip()
                    match = re.search(r"[-+]?\d*\.\d+|\d+", line)
                    if match:
                        current_value = float(match.group())
                        self.data.append(current_value)
                        self.corriente.config(text=f"Corriente = {current_value}")

                head = self.ser.read(1)
                if head != b'\x55':
                    continue

                head = self.ser.read(1)
                if head == b'\x61':
                    a = np.round(np.array(struct.unpack('hhh', self.ser.read(6))) / 32768 * 16, 3)
                    self.aa = np.vstack((self.aa, a))
                    elapsed_time = time.time() - self.start_time
                    self.tt = np.append(self.tt, elapsed_time)

                    self.contador += 1

                    if self.contador == 70:
                        self.arduino.write(b'ON\n')
                        print("MOTOR ENCENDIDO")

                    if self.contador == 400:
                        self.arduino.write(b'OFF\n')
                        print("MOTOR APAGADO")

                    if self.contador == 500:
                        self.running = False
                        self.save_data()

                self.ser.read(3)
        except Exception as e:
            messagebox.showerror("Error", f"Error in data collection: {e}")
            self.stop_data_collection()

    def update_graph(self):
        if self.running:
            if len(self.tt) > 15:
                tt_to_plot = self.tt[15:]
                aa_to_plot = self.aa[15:]

                aa_to_plot_centered = aa_to_plot - np.mean(aa_to_plot, axis=0)
                fft_x = np.fft.fft(aa_to_plot_centered[:, 0])
                fft_y = np.fft.fft(aa_to_plot_centered[:, 1])
                fft_z = np.fft.fft(aa_to_plot_centered[:, 2])

                freqs = np.fft.fftfreq(len(aa_to_plot), d=0.01)

                for i, fft_line in enumerate(self.fft_lines):
                    fft_line.set_data(freqs[:len(fft_x)//2], np.abs([fft_x, fft_y, fft_z][i])[:len(fft_x)//2])
                    self.axs[i, 1].relim()
                    self.axs[i, 1].autoscale_view()

                for i, line in enumerate(self.lines):
                    line.set_data(tt_to_plot, aa_to_plot[:, i])
                    self.axs[i, 0].relim()
                    self.axs[i, 0].autoscale_view()

                self.canvas.draw()

            self.master.after(100, self.update_graph)

if __name__ == "__main__":
    root = Tk()
    app = AccelerometerGUI(root)
    root.mainloop()