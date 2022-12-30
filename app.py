import tkinter as tk
import tkinter.filedialog as filedialog
import customtkinter
import excel_driver
import threading


filepaths = {
    'data_filepath': '',
    'report_filepath': '' 
}

class Root(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("pyBookingReport")
        self.geometry("720x420")

        self.main_frame = customtkinter.CTkFrame(master=self) 
        self.main_frame.pack(padx = 20, pady= 30, expand=False, fill=tk.BOTH)

        self.data_button = customtkinter.CTkButton(master=self.main_frame, command=self.open_data_file, text = "Data File")
        self.data_button.pack(pady=25, padx=50)

        self.invalid_label_1 = customtkinter.CTkLabel(master=self.main_frame, text="")
    
        self.report_button = customtkinter.CTkButton(master=self.main_frame, command=self.open_report_file, text = "Booking Report")
        self.report_button.pack(pady=25, padx= 50) 

        self.submit_button = customtkinter.CTkButton(master=self, command=self.on_submit, text = "Submit")
        self.submit_button.pack(pady = 5, padx = 0)

        self.progress_bar = customtkinter.CTkProgressBar(master=self)
        self.progress_bar.pack(pady = 15)
        self.progress_bar.set(0)
        
       

        self.progress_label = customtkinter.CTkLabel(master=self, text = "")
        self.progress_label.pack()

    def open_data_file(self):
        # Open an "open file" dialog and get the selected file's path
        file_path = filedialog.askopenfilename()
        filepaths['data_filepath'] = file_path

        if file_path != "":
            path_arr = file_path.split(".")
            file_name = file_path.split("/") 
            file_name = file_name[len(file_name)-1] 
            if path_arr[1] != "xlsx":
                self.invalid_label_1.configure(text = f"{file_name} is not of file extension .xlsx")
                self.invalid_label_1.pack() 
            else:
                self.invalid_label_1.pack_forget() 
        print("data")

    def open_report_file(self):
        file_path = filedialog.askopenfilename()
        filepaths['report_filepath'] = file_path
        print("report")

    def on_submit(self):
        if filepaths['data_filepath'] != '' and filepaths['report_filepath'] != '':
            thread_1 = threading.Thread(target=excel_driver.driver, args=(filepaths,self.progress_label,self.progress_bar,))
            thread_1.start() 
            #excel_driver.driver(filepaths)
            filepaths['report_filepath'] = '' 
            filepaths['data_fileplath'] = ''


if __name__ == "__main__":
    app = Root()
    app.mainloop()

