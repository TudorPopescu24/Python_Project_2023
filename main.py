import tkinter as tk
import matplotlib
import psutil
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from tkinter import Menu
import sqlite3
import datetime
from datetime import datetime
from tkcalendar import Calendar
from xlsxwriter.workbook import Workbook

# Chart styling
plt.style.use('ggplot')
matplotlib.rcParams['lines.linewidth'] = 1


class SystemMonitor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.conn = sqlite3.connect('resourceMonitor.db')

        self.title('GUI Resource Monitor')

        menu = Menu(self)
        self.config(menu=menu)
        menu.add_command(label='Export data', command=lambda: show_popup(self.conn))

        # Seconds counter for charts
        self.secondCount = 60

        # CPU Widgets
        self.cpu_title_label = tk.Label(self, text="Information about CPU", font="Arial 18 bold")
        self.cpu_usage_graph = tk.Canvas(self, width=400, height=50, bg="white")
        self.cpu_usage_label = tk.Label(self, text="CPU Usage:", font=("Arial", 12))
        self.cpu_export_button = tk.Button(self, text="Export Chart", command=self.export_cpu_chart)
        self.cpu_title_label.grid(row=0, column=0, columnspan=2)
        self.cpu_usage_graph.grid(row=1, column=0, padx=10)
        self.cpu_usage_label.grid(row=1, column=1)
        self.cpu_export_button.grid(row=1, column=3)

        self.cpu_figure = Figure(figsize=(10, 3), dpi=80)
        self.cpu_plot = self.cpu_figure.add_subplot()

        self.cpu_plot_xList = []
        self.cpu_plot_yList = []

        self.cpu_graph_canvas = FigureCanvasTkAgg(self.cpu_figure, self)
        self.cpu_graph_canvas.get_tk_widget().grid(row=0, column=2, padx=10, pady=10, rowspan=3)
        self.cpu_plot.plot(self.cpu_plot_xList, self.cpu_plot_yList, color="blue")
        self.cpu_plot.set_title('CPU Usage')
        self.cpu_plot.invert_xaxis()

        # Memory Usage Widgets
        self.memory_title_label = tk.Label(self, text="Information about Memory & Disk", font="Arial 18 bold")
        self.memory_usage_graph = tk.Canvas(self, width=400, height=50, bg="white")
        self.memory_usage_label = tk.Label(self, text="Memory Usage:", font=("Arial", 12))
        self.memory_export_button = tk.Button(self, text="Export Chart", command=self.export_memory_chart)
        self.memory_title_label.grid(row=3, column=0, columnspan=2)
        self.memory_usage_graph.grid(row=4, column=0)
        self.memory_usage_label.grid(row=4, column=1)
        self.memory_export_button.grid(row=4, column=3)

        self.memory_figure = Figure(figsize=(10, 3), dpi=80)
        self.memory_plot = self.memory_figure.add_subplot()

        self.memory_plot_xList = []
        self.memory_plot_yList = []

        self.memory_graph_canvas = FigureCanvasTkAgg(self.memory_figure, self)
        self.memory_graph_canvas.get_tk_widget().grid(row=3, column=2, padx=10, pady=10, rowspan=3)
        self.memory_plot.plot(self.memory_plot_xList, self.memory_plot_yList, color="green")
        self.memory_plot.set_title('RAM Usage')
        self.memory_plot.invert_xaxis()

        # Disk Space Widgets
        self.disk_space_graph = tk.Canvas(self, width=400, height=50, bg="white")
        self.disk_space_label = tk.Label(self, text="Disk Space:", font=("Arial", 12))
        self.disk_space_graph.grid(row=5, column=0)
        self.disk_space_label.grid(row=5, column=1)

        # Network Usage Widgets
        self.memory_title_label = tk.Label(self, text="Information about Network", font="Arial 18 bold")
        self.bytes_sent_label = tk.Label(self, text="Send:", font=("Arial", 14))
        self.bytes_recv_label = tk.Label(self, text="Receive:", font=("Arial", 14))
        self.network_export_button = tk.Button(self, text="Export Chart", command=self.export_network_chart)
        self.memory_title_label.grid(row=6, column=0, columnspan=2)
        self.bytes_sent_label.grid(row=7, column=0, columnspan=2)
        self.bytes_recv_label.grid(row=8, column=0, columnspan=2)
        self.network_export_button.grid(row=7, column=3)

        self.network_figure = Figure(figsize=(10, 3), dpi=80)
        self.network_plot = self.network_figure.add_subplot()

        self.network_plot_xList = []
        self.network_plot_send_yList = []
        self.network_plot_recv_yList = []

        self.network_graph_canvas = FigureCanvasTkAgg(self.network_figure, self)
        self.network_graph_canvas.get_tk_widget().grid(row=6, column=2, padx=10, pady=10, rowspan=3)
        self.network_plot.plot(self.network_plot_xList, self.network_plot_send_yList, color="red")
        self.network_plot.plot(self.network_plot_xList, self.network_plot_recv_yList, color="orange")
        self.network_plot.set_title('Network Usage')
        self.network_plot.invert_xaxis()

        net_info = psutil.net_io_counters()
        self.prev_bytes_sent = net_info.bytes_sent / 1024
        self.prev_bytes_recv = net_info.bytes_recv / 1024

        # Update display
        self.after(1000, self.update_display)

    def update_display(self):
        # Update widgets
        self.update_cpu_widgets()
        self.update_memory_widgets()
        self.update_disk_widgets()
        self.update_network_widgets()

        # 1 second has passed
        self.secondCount -= 1

        # Update display every second
        self.after(1000, self.update_display)

    def update_cpu_widgets(self):
        # Get CPU usage
        cpu_usage = psutil.cpu_percent()

        # Update CPU usage graph
        self.cpu_usage_graph.delete("all")
        self.cpu_usage_graph.create_rectangle(0, 0, 400 * (cpu_usage / 100), 200, fill="blue")

        # Update CPU usage label
        self.cpu_usage_label.config(text="CPU Usage: {:.2f}%".format(cpu_usage), font=("Arial", 12))

        # Update CPU chart
        if self.secondCount > 1:
            self.cpu_plot_xList.append(self.secondCount)
            self.cpu_plot_yList.append(cpu_usage)
        else:
            self.cpu_plot_yList.pop(0)
            self.cpu_plot_yList.append(cpu_usage)

        self.cpu_plot.cla()
        self.cpu_plot.plot(self.cpu_plot_xList, self.cpu_plot_yList, color="blue")
        self.cpu_plot.set_title('CPU Usage')
        self.cpu_plot.invert_xaxis()
        self.cpu_graph_canvas.draw()

        # Insert data into the table
        cursor = self.conn.cursor()
        cursor.execute("INSERT INTO CPU (Value, ReadDate) VALUES (?, ?)",
                       (cpu_usage, datetime.now().strftime("%Y-%m-%d-%H:%M:%S")))
        self.conn.commit()

    def update_memory_widgets(self):
        # Get memory usage
        memory_usage = psutil.virtual_memory().percent

        # Update memory usage bar
        self.memory_usage_graph.delete("all")
        self.memory_usage_graph.create_rectangle(0, 0, 400 * (memory_usage / 100), 200, fill="green")

        # Update memory usage label
        self.memory_usage_label.config(text="Memory Usage: {:.2f}%".format(memory_usage), font=("Arial", 12))

        # Update memory chart
        if self.secondCount > 1:
            self.memory_plot_xList.append(self.secondCount)
            self.memory_plot_yList.append(memory_usage)
        else:
            self.memory_plot_yList.pop(0)
            self.memory_plot_yList.append(memory_usage)

        self.memory_plot.cla()
        self.memory_plot.plot(self.memory_plot_xList, self.memory_plot_yList, color="green")
        self.memory_plot.set_title('RAM Usage')
        self.memory_plot.invert_xaxis()
        self.memory_graph_canvas.draw()

        # Insert data into the table
        cursor = self.conn.cursor()
        cursor.execute("INSERT INTO Memory (Value, ReadDate) VALUES (?, ?)",
                       (memory_usage, datetime.now().strftime("%Y-%m-%d-%H:%M:%S")))
        self.conn.commit()

    def update_disk_widgets(self):
        # Get disk space usage
        disk_usage = psutil.disk_usage("/").percent

        # Update disk space bar
        self.disk_space_graph.delete("all")
        self.disk_space_graph.create_rectangle(0, 0, 400 * (disk_usage / 100), 200, fill="yellow")

        # Update disk space label
        self.disk_space_label.config(text="Disk Space: {:.2f}%".format(disk_usage), font=("Arial", 12))

    def update_network_widgets(self):
        # Get network statistics
        net_info = psutil.net_io_counters()
        send_data = net_info.bytes_sent / 1024
        recv_data = net_info.bytes_recv / 1024
        send_kbps = send_data - self.prev_bytes_sent
        recv_kbps = recv_data - self.prev_bytes_recv
        self.bytes_sent_label.config(text="Send: {:.2f} KBps".format(send_kbps), font=("Arial", 14))
        self.bytes_recv_label.config(text="Receive: {:.2f} KBps".format(recv_kbps), font=("Arial", 14))
        self.prev_bytes_sent = send_data
        self.prev_bytes_recv = recv_data

        # Update network chart
        if self.secondCount > 1:
            self.network_plot_xList.append(self.secondCount)
            self.network_plot_send_yList.append(send_kbps)
            self.network_plot_recv_yList.append(recv_kbps)
        else:
            self.network_plot_send_yList.pop(0)
            self.network_plot_recv_yList.pop(0)
            self.network_plot_send_yList.append(send_kbps)
            self.network_plot_recv_yList.append(recv_kbps)

        self.network_plot.cla()
        self.network_plot.plot(self.network_plot_xList, self.network_plot_send_yList, color="red")
        self.network_plot.plot(self.network_plot_xList, self.network_plot_recv_yList, color="orange")
        self.network_plot.set_title('Network Usage')
        self.network_plot.invert_xaxis()
        self.network_graph_canvas.draw()

        # Insert data into the table
        cursor = self.conn.cursor()
        cursor.execute("INSERT INTO Network (SendValue, RecvValue, ReadDate) VALUES (?, ?, ?)",
                       (send_kbps, recv_kbps, datetime.now().strftime("%Y-%m-%d-%H:%M:%S")))
        self.conn.commit()

    def export_cpu_chart(self):
        self.cpu_figure.savefig("cpu_chart.jpg")

    def export_memory_chart(self):
        self.memory_figure.savefig("memory_chart.jpg")

    def export_network_chart(self):
        self.network_figure.savefig("network_chart.jpg")


def show_popup(conn):
    # Create the popup window
    popup = tk.Tk()
    popup.title("Export data")

    # Start date
    start_date_label = tk.Label(popup, text="Select start date:", font=("Arial", 12))
    start_date = Calendar(popup, selectmode="day", year=2023, month=1, date=1, date_pattern="y-mm-dd-00:00:00")
    start_date_label.grid(row=0, column=0)
    start_date.grid(row=1, column=0, padx=10)

    # End date
    end_date_label = tk.Label(popup, text="Select end date:", font=("Arial", 12))
    end_date = Calendar(popup, selectmode="day", year=2023, month=1, date=1, date_pattern="y-mm-dd-00:00:00")
    end_date_label.grid(row=0, column=2)
    end_date.grid(row=1, column=2, padx=10)

    # Export data button
    button = tk.Button(popup, text="Export data",
                       command=lambda: export_data(start_date.get_date(), end_date.get_date(), conn), bg="gray",
                       fg="black")
    button.grid(row=2, column=1, pady=10)

    # Run the popup loop
    popup.mainloop()


def export_data(start_date, end_date, conn):
    # Create Excel file
    workbook = Workbook('data_output.xlsx')

    # Export data to Excel file
    export_data_cpu(conn, workbook, start_date, end_date)
    export_data_memory(conn, workbook, start_date, end_date)
    export_data_network(conn, workbook, start_date, end_date)

    # Close Excel file
    workbook.close()


def export_data_cpu(conn, workbook, start_date, end_date):
    bold = workbook.add_format({'bold': True})
    worksheet = workbook.add_worksheet('CPU')

    cursor = conn.cursor()
    cursor.execute("SELECT Value, ReadDate FROM CPU WHERE ReadDate >= ? AND ReadDate <= ?", (start_date, end_date))

    rows = cursor.fetchall()

    worksheet.write(0, 0, "Value (%)", bold)
    worksheet.write(0, 1, "Date", bold)
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            worksheet.write(i + 1, j, row[j])


def export_data_memory(conn, workbook, start_date, end_date):
    bold = workbook.add_format({'bold': True})
    worksheet = workbook.add_worksheet('Memory')

    cursor = conn.cursor()
    cursor.execute("SELECT Value, ReadDate FROM Memory WHERE ReadDate >= ? AND ReadDate <= ?", (start_date, end_date))

    rows = cursor.fetchall()

    worksheet.write(0, 0, "Value (%)", bold)
    worksheet.write(0, 1, "Date", bold)
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            worksheet.write(i + 1, j, row[j])


def export_data_network(conn, workbook, start_date, end_date):
    bold = workbook.add_format({'bold': True})
    worksheet = workbook.add_worksheet('Network')

    cursor = conn.cursor()
    cursor.execute("SELECT SendValue, RecvValue, ReadDate FROM Network WHERE ReadDate >= ? AND ReadDate <= ?",
                   (start_date, end_date))

    rows = cursor.fetchall()

    worksheet.write(0, 0, "Send (KBps)", bold)
    worksheet.write(0, 1, "Receive (KBps)", bold)
    worksheet.write(0, 2, "Date", bold)
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            worksheet.write(i + 1, j, row[j])


if __name__ == "__main__":
    app = SystemMonitor()
    app.mainloop()
    app.conn.close()
