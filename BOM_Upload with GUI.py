
from tkinter import *
import tkinter as tk, pandas as pd
import tkinter.messagebox, tkinter.filedialog, xmlrpc.client, glob, os

class mainWindow():

    def __init__(self, Window):  # constructor

        self.Window = Window
        self.Window.title(" Odoo BoM Import Tool V.1.1 ")
        self.Window.geometry("800x400")

        self.LoginForm = tk.Frame(Window)
        self.LoginForm.pack()

        self.TitleLabel = tk.Label(self.LoginForm, text="Import Tool", font='Helvetica 8 bold')
        self.TitleLabel.grid(row=2, column=1)

        self.URLLabel = tk.Label(self.LoginForm, text="Odoo URL:")
        self.URLLabel.grid(row=4, column=0)
        self.URLEntry = tk.Entry(self.LoginForm)
        self.URLEntry.grid(row=4, column=1)
        self.DBLabel = tk.Label(self.LoginForm, text="Database:")
        self.DBLabel.grid(row=5, column=0)
        self.DBEntry = tk.Entry(self.LoginForm)
        self.DBEntry.grid(row=5, column=1)
        self.UserLabel = tk.Label(self.LoginForm, text="Username:")
        self.UserLabel.grid(row=6, column=0)
        self.UserEntry = tk.Entry(self.LoginForm)
        self.UserEntry.grid(row=6, column=1)
        self.PasswordLabel = tk.Label(self.LoginForm, text="Password:")
        self.PasswordLabel.grid(row=7, column=0)
        self.PasswordEntry = tk.Entry(self.LoginForm, show="*")
        self.PasswordEntry.grid(row=7, column=1)

        self.LoginButton = tk.Button(self.LoginForm, text="Log in to Odoo", command=self.action_login)
        self.LoginButton.grid(row=9, column=1)

        self.TopLevelAssyLabel = tk.Label(self.LoginForm, text="Top-Level Assembly Information", font='Helvetica 8 bold')
        self.TopLevelAssyLabel.grid(row=10, column=1)
        self.TLAssyPNLabel = tk.Label(self.LoginForm, text="Module Number:")
        self.TLAssyPNLabel.grid(row=11, column=0)
        self.TLAssyPNEntry = tk.Entry(self.LoginForm)
        self.TLAssyPNEntry.grid(row=11, column=1)
        self.TLAssyNameLabel = tk.Label(self.LoginForm, text="Module Name:")
        self.TLAssyNameLabel.grid(row=12, column=0)
        self.TLAssyNameEntry = tk.Entry(self.LoginForm)
        self.TLAssyNameEntry.grid(row=12, column=1)
        self.TLAssyRevLabel = tk.Label(self.LoginForm, text="Module Revision:")
        self.TLAssyRevLabel.grid(row=13, column=0)
        self.TLAssyRevEntry = tk.Entry(self.LoginForm)
        self.TLAssyRevEntry.grid(row=13, column=1)

        self.BOMLabel = tk.Label(self.LoginForm, text="Bills of Material", font='Helvetica 8 bold')
        self.BOMLabel.grid(row=15, column=1)
        self.FileLabel = tk.Label(self.LoginForm, text="File:")
        self.FileLabel.grid(row=16, column=0)
        self.FileEntry = tk.Label(self.LoginForm, wraplength=120)
        self.FileEntry.grid(row=16, column=1)
        self.FileButton = tk.Button(self.LoginForm, text="Browse...", command=self.browseFiles)
        self.FileButton.grid(row=16, column=2)
        self.PartsOnlyEntry = tk.Checkbutton(self.LoginForm, text="Import Components Only", variable=upload_parts_only)
        self.PartsOnlyEntry.grid(row=17, column=1)
        self.NewBOMRevEntry = tk.Checkbutton(self.LoginForm, text="Create New BOM Revision", variable=create_new_BOM_revision)
        self.NewBOMRevEntry.grid(row=18, column=1)
        self.ImportButton = tk.Button(self.LoginForm, text="Import", command=self.action_import, state='disabled')
        self.ImportButton.grid(row=19, column=1)

        # Coming Soon!
        self.NewBOMRevEntry.config(state='disabled')

        try:
            with open("credentials.txt", "r") as filestream:
                for line in filestream:
                    currentline = line.split(",")
                    self.URLEntry.insert(0, currentline[0])
                    self.DBEntry.insert(0, currentline[1])
                    self.UserEntry.insert(0, currentline[2])
                    self.PasswordEntry.insert(0, currentline[3])
        except:
            consoleList.insert(END, 'Either credential file is missing, or information is incomplete.')


    # File Browser
    def browseFiles(self):
        self.filename = tkinter.filedialog.askopenfilename(initialdir="/", title="Select a File", filetypes=[('CSV', '*.csv',), ('Excel', ('*.xls', '*.xlsx'))])
        self.FileEntry.configure(text=self.filename)

    # Unlock Text Inputs
    def unlock_credentials(self):
        self.URLEntry.config(state='normal')
        self.DBEntry.config(state='normal')
        self.UserEntry.config(state='normal')
        self.PasswordEntry.config(state='normal')

    # Read BOM and sanitize BOM fields (remove indentation and null values)
    def readBOM(self, filename):
        BOM = pd.read_csv(filename, encoding='unicode_escape')
        BOM['Level'] = BOM['Level'].astype(str)
        #BOM['Level Computed'] = BOM['Level'].str.count("\\.")

        BOM['PART NUMBER'] = BOM['PART NUMBER'].str.strip()
        BOM['REVISION'].fillna('0', inplace=True)
        BOM.fillna('', inplace=True)
        BOM['REVISION'] = BOM['REVISION'].astype(str)
        BOM['QTY.'] = BOM['QTY.'].astype(int)

        return BOM

    def findParentAssy(self, BOM, i):
        level = BOM['Level'][i]
        parent = {}

        # if level value has no periods, assign top level assembly as the parent assembly
        if (level.count('.') == 0):
            parent['assynumber'] = toplvl_assy
            parent['rev'] = toplvl_rev

        # else, retrieve level of parent assy (ex: 1.2.3 -> 1.2)
        else:
            # split level value by the last occurence of "."
            parent_lvl = level.rpartition('.')[0]
            try:
                # get row containing parent assy
                parentBOM = BOM.loc[BOM['Level'] == parent_lvl]
                # get part number & rev
                parent['assynumber'] = parentBOM.iloc[0, 1]
                parent['rev'] = parentBOM.iloc[0, 2]
            except:
                print('Issue finding parent level:', parent_lvl)
                consoleList.insert(END, 'ERROR: Cannot locate Parent Level!')

        return parent

    def uploadBOM(self, BOM, uid):

        self.editedBOMlist = []

        # Verify Top-Level Assembly exists in Odoo
        toplvl_assy_id = self.searchProduct(uid, toplvl_assy, toplvl_rev)
        if not toplvl_assy_id:
            self.createProduct(uid, toplvl_assy, toplvl_rev, toplvl_name, '', '', '', '', '', '','', '')

        for i in range(BOM.shape[0]):

            eng_code = BOM['PART NUMBER'][i]
            print('Processing', eng_code)

            # Clean revision notations of P items and items without revisions.
            if (eng_code[0]=='P' and eng_code[1].isdigit()):
                rev = '0'
                # BOM['REVISION'][i] = '0'
            else:
                rev = BOM['REVISION'][i]

            desc = BOM['DESCRIPTION'][i]
            qty = BOM['QTY.'][i]
            material = BOM['BOMMaterial'][i]
            length = BOM['BOM LENGTH'][i]
            finish = BOM['FINISH'][i]
            finspec = BOM['FINISH SPEC'][i]
            vendor = BOM['VENDOR'][i]
            vendorno = BOM['VENDORNO'][i]
            spareclass = BOM['SPARECLASS'][i]
            subclass = BOM['SUBSTITUTION CLASS'][i]

            # Check if product exists in Odoo.
            prod_id = self.searchProduct(uid, eng_code, rev)

            # Create products if it does not exist in DB
            if len(prod_id) == 0:
                try:
                    self.createProduct(uid,eng_code,rev,desc,material,length,finish,finspec,vendor,vendorno,spareclass,subclass)
                    consoleList.insert(END, eng_code + ': Component created.')
                except:
                    # if exception, try to edit product instead
                    try:
                        self.editProduct(uid, prod_id, desc, material, length, finish, finspec, vendor, vendorno, spareclass, subclass)
                        consoleList.insert(END, eng_code + ': Component updated after failing to create.')
                    except:
                        print('Update failed.')
                        consoleList.insert(END, 'ERROR: Update failed when trying to update ' + eng_code + '!')

            # Update products if already exist
            else:
                try:
                    self.editProduct(uid, prod_id, desc, material, length, finish, finspec, vendor, vendorno, spareclass, subclass)
                    consoleList.insert(END, eng_code + ': Component updated.')
                except:
                    consoleList.insert(END, 'ERROR: Update failed when trying to update ' + eng_code + '!')

            # Create Bill of Material.
            if not upload_parts_only.get():

                parent_assy = self.findParentAssy(BOM, i)['assynumber']
                parent_assy_rev = self.findParentAssy(BOM, i)['rev']

                # Check for recursive references
                if parent_assy and parent_assy != eng_code:
                    parent_assy_id = self.searchProduct(uid, parent_assy, parent_assy_rev)
                    parent_bom = self.searchLatestBOM(uid, parent_assy_id)

                    # Scenario 1: Parent BOM exists, update current BOM
                    if parent_bom['id'] and create_new_BOM_revision.get() != 1:
                        # delete all pre-existing BOM Lines if not already done in this session
                        if parent_bom['id'] not in editedBOMlist:
                            self.deleteBOMLine(uid, parent_bom['id'])
                        self.editBOMLine(uid, [parent_bom['id']], eng_code, rev, qty)
                        consoleList.insert(END, eng_code + ': Added to BOM of ' + parent_assy + '.')
                    # Scenario 2: Parent BOM exists, create new BOM revision
                    elif parent_bom['id'] and create_new_BOM_revision.get() == 1:
                        if parent_bom['id'] not in self.editedBOMlist:
                            print(parent_bom['id'])
                            print(self.editedBOMlist)
                            print('Not in List')
                            parent_bom['rev']+=1
                            parent_bom_id = [self.createBOM(uid, parent_assy_id[0], parent_bom['rev'])]
                        self.editBOMLine(uid, parent_bom_id, eng_code, rev, qty)
                        consoleList.insert(END, eng_code + ': Added to BOM of ' + parent_assy + '.')
                    # Scenario 3: Parent BOM does not exist, create new BOM revision
                    else:
                        parent_bom_id = [self.createBOM(uid, parent_assy_id[0], parent_bom['rev'])]
                        self.editBOMLine(uid, parent_bom_id, eng_code, rev, qty)
                        consoleList.insert(END, eng_code + ': Added to BOM of ' + parent_assy + '.')

        consoleList.insert(END, 'Upload completed.')

    def searchProduct(self, uid, eng_code, rev):
        prod_id = models.execute_kw(db, uid, password, 'product.template', 'search',
                                    [[['engineering_code', '=', eng_code],
                                      ['engineering_revision_text', '=', rev]]])
        return prod_id

    def createProduct(self, uid, eng_code, rev, desc, material, length, finish, finspec, vendor, vendorno, spareclass, subclass):
        models.execute_kw(db, uid, password, 'product.product', 'create',
                          [{'engineering_code': eng_code, 'engineering_revision_text': rev,
                            'name': desc, 'engineering_material': material, 'profile_spec': length, 'engineering_surface': finish,
                            'surface_spec': finspec, 'vendor_name': vendor, 'vendor_prod_code': vendorno,
                            'spare_class': spareclass, 'substituion_class': subclass}])

    def editProduct(self, uid, prod_id, desc, material, length, finish, finspec, vendor, vendorno, spareclass, subclass):
        models.execute_kw(db, uid, password, 'product.template', 'write',
                          [prod_id, {'name': desc, 'engineering_material': material, 'profile_spec': length, 'engineering_surface': finish,
                                     'surface_spec': finspec, 'vendor_name': vendor, 'vendor_prod_code': vendorno,
                                     'spare_class': spareclass, 'substituion_class': subclass}])

    def searchLatestBOM(self, uid, parent_assy_id):
        bom_ids = models.execute_kw(db, uid, password, 'mrp.bom', 'search_read',
                                    [[['product_tmpl_id', '=', parent_assy_id]]],
                                    {'fields': ['engineering_revision']})

        if bom_ids:
            bom_ids = pd.DataFrame(bom_ids)
            max_rev = bom_ids['engineering_revision'].idxmax()
            bom_id = int(bom_ids['id'][max_rev])

            return {'id': bom_id, 'rev': max_rev}
        else:
            return {'id': False, 'rev': 0}

    def createBOM(self, uid, parent_assy_id, rev):
        bom_id = models.execute_kw(db, uid, password, 'mrp.bom', 'create',
                                   [{'product_tmpl_id': parent_assy_id, 'engineering_revision': rev}])
        editedBOMlist.append(bom_id)
        print('Appended:', bom_id, 'in', self.editedBOMlist)
        return bom_id

    def createBOMLine(self, uid, bom_id, prod_id, qty):
        models.execute_kw(db, uid, password, 'mrp.bom.line', 'create',
                          [{'product_id': prod_id, 'product_qty': qty, 'bom_id': bom_id}])

    def editBOM(self, uid, prod_id, desc):
        models.execute_kw(db, uid, password, 'mrp.bom', 'write',
                          [prod_id, {'name': desc}])

    def searchBOMItem(self, uid, eng_code, rev):
        prod_exists = models.execute_kw(db, uid, password, 'mrp.bom', 'search',
                                        [[['product_tmpl_id', '=', eng_code],
                                          ['engineering_revision_text', '=', rev]]])
        return prod_exists

    def editBOMLine(self, uid, bom_id, eng_code, rev, qty):
        # Convert qty from int32 to int
        qty = int(qty)
        # find product.product item
        product_product_id = models.execute_kw(db, uid, password, 'product.product', 'search',
                                               [[['engineering_code', '=', eng_code],
                                                 ['engineering_revision_text', '=', rev]]])
        # find BOM line id
        bom_line_id = models.execute_kw(db, uid, password, 'mrp.bom.line', 'search',
                                        [[['bom_id', '=', bom_id],
                                          ['product_id', '=', product_product_id]]])
        # write onto BOM Line
        if bom_line_id:
            models.execute_kw(db, uid, password, 'mrp.bom.line', 'write',
                              [bom_line_id, {'product_qty': qty}])
        else:
            models.execute_kw(db, uid, password, 'mrp.bom.line', 'create',
                              [{'product_id': int(product_product_id[0]), 'product_qty': qty,
                                'bom_id': int(bom_id[0])}])

    def deleteBOMLine(self, uid, bom_id):
        bom_line_ids = models.execute_kw(db, uid, password, 'mrp.bom.line', 'search',
                                         [[['bom_id', '=', bom_id]]])
        models.execute_kw(db, uid, password, 'mrp.bom.line', 'unlink', [bom_line_ids])
        # Avoid erasing the same BOM more than once
        editedBOMlist.append(bom_id)

    # Log in Odoo
    def action_login(self):
        url = self.URLEntry.get()
        global db
        db = self.DBEntry.get()
        username = self.UserEntry.get()
        global password
        password = self.PasswordEntry.get()

        if url and db and username and password:
            self.URLEntry.config(state='disabled')
            self.DBEntry.config(state='disabled')
            self.UserEntry.config(state='disabled')
            self.PasswordEntry.config(state='disabled')
            try:
                global common
                common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
                global models
                models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
                global uid
                uid = common.authenticate(db, username, password, {})
                if uid:
                    self.LoginButton['text'] = 'Log in successful'
                    consoleList.insert(END, 'Log in successful as ' + username + '.')
                    self.LoginButton.config(state='disabled')
                    self.ImportButton.config(state='normal')
                    consoleList.insert(END, 'Ready to Import...')
                else:
                    self.unlock_credentials()
                    tkinter.messagebox.showinfo("Error", "User not found.")
            except Exception as e:
                self.unlock_credentials()
                tkinter.messagebox.showinfo("Error", "Log in failed.")
                if hasattr(e, 'message'):
                    print(e.message)
                else:
                    print(e)
        else:
            self.unlock_credentials()
            tkinter.messagebox.showinfo("Error", "Missing credentials!")

    # Import BOM to Odoo
    def action_import(self):
        filename = self.FileEntry['text']
        if self.TLAssyNameEntry and self.TLAssyRevEntry and self.TLAssyPNEntry and filename:
            global toplvl_assy
            toplvl_assy = self.TLAssyPNEntry.get()
            global toplvl_rev
            toplvl_rev = self.TLAssyRevEntry.get()
            global toplvl_name
            toplvl_name = self.TLAssyNameEntry.get()
            BOM = self.readBOM(filename)
            self.uploadBOM(BOM, uid)
        else:
            tkinter.messagebox.showinfo("Error", "You are missing some fields. Please check again.")

# Keep window in center
def centerWindow(width=300, height=200):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    root.geometry('%dx%d+%d+%d' % (width, height, x, y))

root = tk.Tk()
centerWindow(400, 400)
create_new_BOM_revision = tk.IntVar()
upload_parts_only = tk.IntVar()
editedBOMlist = []
root.resizable(False, False)
scrollbar = Scrollbar(root)
scrollbar.pack(side=RIGHT, fill=Y)
consoleList = Listbox(root, yscrollcommand=scrollbar.set, bg='Black', fg='Green', height='20')
consoleList.insert(END, 'System Ready, waiting for credentials...')
consoleList.pack(side=RIGHT, fill='both', expand=True)
scrollbar.config(command=consoleList.yview)
mainWindow(root)
root.mainloop()
