#!/usr/bin/python

# Author: Saicharan.G
# Code: Email automation

try:
      from tkinter import *
      from tkinter import filedialog
      from tkinter import messagebox
      from threading import Thread
      import smtplib
      import xlrd
      import xlwt
      import os
except ImportError:
      print( '\n>>> You dont\'ve required modules, please install all modules\n' )
      quit()

def show_message( message, die = False ):
      messagebox.showinfo( 'Information', message )
      if die:
            exit( 1 )

file = label = thread = window = None

def update_status( status ):
      label.config( text = 'Status :     ' + status )
      window.update()

class Mail:
      smtp = None
      sender_mail = None
      sender_mail_password = None
      logged_in = None

      def __init__( self, sender_mail, sender_mail_password ):
            self.sender_mail = sender_mail
            self.sender_mail_password = sender_mail_password
            self.logged_in = False
            update_status( 'Connecting to SMTP server...' )
            try:
                  self.smtp = smtplib.SMTP( 'smtp.gmail.com', 587 )
            except Exception as exception:
                  show_message( 'We got an error while connecting Email server, try agin' )
                  return
            self.smtp.ehlo()
            self.smtp.starttls()
            try:
                  self.smtp.login( sender_mail, sender_mail_password )
                  self.logged_in = True
            except smtplib.SMTPAuthenticationError:
                  show_message( 'We got an error while log in.\nCheck your credentials...\nTurn on less secure apps option for your mail' )
            except smtplib.SMTPConnectError:
                  show_message( 'We got an error while connecting' )
            except Exception as exception:
                  show_message( exception.message )

      def send_mail( self, reciever_mail, message, subject = 'Alert' ):
            try:
                  self.smtp.sendmail( self.sender_mail, reciever_mail, 'Subject:{0}\n{1}'.format(subject, message) )
            except:
                  return True

      def __del__( self ):
            if self.smtp:
                  self.smtp.close()
            update_status( 'No process is running' )

def start_process( sender_mail, sender_mail_password ):
      object = Mail( sender_mail, sender_mail_password )
      if not object.logged_in:
            update_status( 'No process is running' )
            return

      try:
            read_book = xlrd.open_workbook( file )
      except:
            show_message( 'We got an error when reading ' + file + '\nCheck that file' )
            return

      write_book = xlwt.Workbook()
      write_sheet = write_book.add_sheet( 'Sheet 1' )
      write_sheet.write( 0, 0, 'Name' )
      write_sheet.write( 0, 1, 'Email' )
      write_sheet.write( 0, 2, 'Candidate ID' )
      write_sheet.write( 0, 3, 'Mobile no' )
      count = 0
      for index in range( len(read_book.sheet_names()) ):
            read_sheet = read_book.sheet_by_index( index )
            count += read_sheet.nrows - 1
      count = str( count )
      update_status( 'Sended mails: 0/' + count )
      for index in range( len(read_book.sheet_names()) ):
            read_sheet = read_book.sheet_by_index( index )
            i = 1
            not_sended = sended = 0
            while i < read_sheet.nrows:
                  try:
                        list = read_sheet.row_values( i )
                        reciever_name, reciever_mail, c_id, reciever_number = list
                        message = 'Hello ' + str(reciever_name) + ',\n\nWe wish to inform you that your payment is due.\nPlease pay using this link.\nYour candidate ID : ' + str(int(c_id)) + '\n\nThanks and Regards.'
                  except Exception as exception:
                        show_message( str(i) + ' line data is incorrect check it.\n' + exception.message, True )
                  if not object.send_mail( reciever_mail, message ):
                        sended += 1
                        update_status( 'Sended mails: ' + str(sended) + '/' + count )
                  else:
                        not_sended += 1
                        for j in range( 4 ):
                              write_sheet.write( not_sended, j, list[j] )
                  i += 1
      final_output = 'Number of users: ' + count + '\nMails sended: ' + str(sended) + '\nMails not sended: ' + str(not_sended)
      if not_sended:
            head, tail = os.path.split( file )
            output = head + '/erros_' + tail
            write_book.save( output )
            final_output += '\n\nFailed mail list created in \"' + output + '\"'
      show_message( final_output )

def main():
      global label, window
      window = Tk()
      window.title( 'Email Automation' )
      window.geometry( '510x170' )
      window.resizable( False, False )
      window.eval( 'tk::PlaceWindow . center' )

      entry1 = Entry( window )
      entry2 = Entry( window )

      class Placeholder:
            text = None

            def __init__( self, text ):
                  self.text = text

            def focus_in( self, event ):
                  if event.widget.get().strip() == self.text:
                        event.widget.delete( 0, 'end' )
                        if event.widget == entry2:
                              event.widget.config( show = '*' )

            def focus_out( self, event ):
                  if not len( event.widget.get().strip() ):
                        event.widget.insert( 0, self.text )
                        if event.widget == entry2:
                              event.widget.config( show = '' )

      entry1.insert( 0, 'Enter your Email' )
      object = Placeholder( 'Enter your Email' )
      entry1.bind( '<FocusIn>', object.focus_in )
      entry1.bind( '<FocusOut>', object.focus_out )
      entry1.grid( row = 0, column = 0, padx = (25, 0), pady = 15, ipady = 4, ipadx = 25 )

      entry2.insert( 0, 'Enter your Email password' )
      object = Placeholder( 'Enter your Email password' )
      entry2.bind( '<FocusIn>', object.focus_in )
      entry2.bind( '<FocusOut>', object.focus_out )
      entry2.grid( row = 0, column = 1, padx = (25, 0), pady = 15, ipady = 4, ipadx = 25 )

      def browse_files( event = None ):
            global file
            file = filedialog.askopenfilename( initialdir = ".", title = "Select .xlsx file", filetypes = (('xlsx files', '*.xlsx'),) )
            entry.insert( 0, file )

      def create_thread():
            global thread
            if not file:
                  show_message( 'Xlsx file is not selected' )
                  return
            if thread and thread.isAlive():
                  show_message( 'Process was already started' )
                  return
            thread = Thread( target = start_process, args = (entry1.get().strip(), entry2.get().strip()) )
            thread.start()

      entry = Entry( window )
      entry.insert( 0, 'Select Xlsx file path' )
      entry.grid( row = 1, column = 0, padx = (25, 0), pady = 15, ipady = 4, ipadx = 25 )
      entry.bind( '<Button-1>', browse_files )

      button1 = Button( window, text = "Select Xlsx file", command = browse_files )
      button1.grid( row = 1, column = 1, padx = 0, pady = 5 )

      label = Label( text = 'Status :     No process is running', font='Helvetica 10 bold' )
      label.grid( row = 2, column = 0, padx = (25, 0), pady = 5 )

      button2 = Button( window, text = "  Send mails   ", command = create_thread )
      button2.grid( row = 2, column = 1, pady = 10 )

      mainloop()

if __name__ == '__main__':
      main()
