# LinkPilot V12 Tree View Edition
import math
import sys, json, os, glob, shutil, threading, webbrowser, base64, tkinter as tk
from tkinter import ttk
from datetime import datetime
from urllib.parse import urlsplit, unquote, quote
from urllib.request import Request, urlopen
from urllib.error import HTTPError


def make_star_icon(master, color, filled):
    size = 16
    center = (size - 1) / 2
    vertices = []
    for point in range(10):
        angle = -math.pi / 2 + point * math.pi / 5
        radius = 7 if point % 2 == 0 else 3
        vertices.append((center + math.cos(angle) * radius,
                         center + math.sin(angle) * radius))
    inner_vertices = [(center + (x - center) * 0.55,
                       center + (y - center) * 0.55) for x, y in vertices]

    def contains(x, y, polygon):
        inside = False
        previous = len(polygon) - 1
        for current in range(len(polygon)):
            x1, y1 = polygon[current]
            x2, y2 = polygon[previous]
            if (y1 > y) != (y2 > y) and x < (x2 - x1) * (y - y1) / (y2 - y1) + x1:
                inside = not inside
            previous = current
        return inside

    image = tk.PhotoImage(master=master, width=size, height=size)
    for y in range(size):
        for x in range(size):
            if contains(x + 0.5, y + 0.5, vertices) and (
                    filled or not contains(x + 0.5, y + 0.5, inner_vertices)):
                image.put(color, (x, y))
    return image


GRAPH_EXPLORER_URL = 'https://developer.microsoft.com/en-us/graph/graph-explorer'


def describe_token(token):
    parts = token.split('.')
    info = f'Token: length={len(token)}, parts={len(parts)}, starts_with_eyJ={token.startswith("eyJ")}'
    try:
        payload = parts[1] + '=' * (-len(parts[1]) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload))
        info += f", aud={claims.get('aud')}, exp={claims.get('exp')}, scp={claims.get('scp')}"
    except Exception as error:
        info += f', payload not decodable ({error})'
    return info


def graph_get(path, token):
    request = Request('https://graph.microsoft.com/v1.0' + path, headers={
        'Authorization': f'Bearer {token}', 'Prefer': 'redeemSharingLinkIfNecessary'})
    try:
        with urlopen(request, timeout=15) as response:
            return json.load(response)
    except HTTPError as error:
        if error.code != 404:
            print(f'Graph {error.code} for {path}: {error.read()[:200]!r}', file=sys.stderr)
        return None


def is_sharepoint_valid(url, token):
    encoded = base64.urlsafe_b64encode(url.encode()).decode().rstrip('=')
    if graph_get(f'/shares/u!{encoded}/driveItem', token) is not None:
        return True

    parsed = urlsplit(url)
    segments = [unquote(part) for part in parsed.path.split('/') if part]
    # Strip the /:x:/r style sharing prefix.
    if segments and segments[0].startswith(':'):
        segments = segments[2:]
    host = parsed.hostname
    lowered = [part.lower() for part in segments]
    cut = lowered.index('_layouts') if '_layouts' in lowered else len(segments)
    for count in range(cut, 1, -1):
        site_path = '/'.join(quote(part) for part in segments[:count])
        site = graph_get(f'/sites/{host}:/{site_path}?$select=id', token)
        if site is None:
            continue
        rest = segments[count:]
        if not rest:
            return True
        drives = graph_get(f"/sites/{site['id']}/drives?$select=id,webUrl", token) or {}
        for drive in drives.get('value', []):
            if unquote(drive['webUrl']).rstrip('/').lower().endswith('/' + rest[0].lower()):
                item_path = rest[1:]
                # Views such as Forms/AllItems.aspx are not drive items.
                if not item_path or item_path[0].lower() == 'forms':
                    return True
                item = '/'.join(quote(part) for part in item_path)
                return graph_get(f"/drives/{drive['id']}/root:/{item}", token) is not None
        return True  # Site is reachable; the sub-path isn't a library we can look up.
    return False


def is_url_valid(url, token=''):
    url = url.strip()
    try:
        parsed = urlsplit(url)
        if parsed.scheme.lower() not in ('http', 'https') or not parsed.netloc:
            return False
        if token and parsed.hostname and parsed.hostname.lower().endswith('sharepoint.com'):
            token = token.strip().strip('"')
            if token.lower().startswith('bearer '):
                token = token[7:].strip()
            return is_sharepoint_valid(url, token)
        request = Request(url, headers={
            'User-Agent': 'LinkPilot URL Validator',
            'Range': 'bytes=0-0'
        })
        with urlopen(request, timeout=15) as response:
            return response.status < 400
    except HTTPError:
        return False
    except Exception as error:
        print(f'Validation error for {url}: {error}', file=sys.stderr)
        return False

if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

import tkinter.messagebox as mb

STATE_FILE=os.path.join(APP_DIR,'linkpilot_state.json')
BACKUP_DIR=os.path.join(APP_DIR,'backup')
os.makedirs(BACKUP_DIR,exist_ok=True)

class TokenDialog(tk.Toplevel):
    def __init__(self, parent, token=''):
        super().__init__(parent)
        self.title('Microsoft Access Token')
        self.resizable(False, False)
        self.transient(parent)
        self.result = None
        link = tk.Label(self, text='Microsoft Access Token', fg='#0563C1', cursor='hand2',
                        font=('Segoe UI', 11, 'underline'))
        link.pack(padx=14, pady=(14, 4), anchor='w')
        link.bind('<Button-1>', lambda e: webbrowser.open(GRAPH_EXPLORER_URL))
        tk.Label(self, text='Optional. Used for SharePoint links.').pack(padx=14, anchor='w')
        self.entry = tk.Entry(self, width=60, show='*')
        self.entry.insert(0, token)
        self.entry.pack(padx=14, pady=8)
        buttons = tk.Frame(self)
        buttons.pack(padx=14, pady=(0, 14), anchor='e')
        tk.Button(buttons, text='Run Validation', command=self.run).pack(side='left', padx=4)
        tk.Button(buttons, text='Cancel validation', command=self.destroy).pack(side='left')
        self.protocol('WM_DELETE_WINDOW', self.destroy)
        self.entry.focus_set()
        self.grab_set()

    def run(self):
        self.result = self.entry.get().strip()
        self.destroy()


class EditDialog(tk.Toplevel):
    def __init__(self,parent,link,save_cb,delete_cb):
        super().__init__(parent)
        self.link=link
        self.save_cb=save_cb
        self.delete_cb=delete_cb
        self.title('Link Details')
        self.transient(parent)
        self.focus_force()
        self.configure(bg='#1e1e1e')
        self.is_new=link.get('new',False)
        description=link['description'].strip()
        show_desc=bool(description) or self.is_new
        w,h=(800,320) if show_desc else (800,180)
        x=parent.winfo_x()+max(0,(parent.winfo_width()-w)//2)
        y=parent.winfo_y()+max(0,(parent.winfo_height()-h)//2)

        self.geometry(f'{w}x{h}+{x}+{y}')
        banner = tk.Frame(
                self,
                bg='#4D148C',
                height=40
        )

        banner.pack(
                fill='x',
                side='top'
        )

        tk.Label(
                banner,
                text='Edit Bookmark',
                bg='#4D148C',
                fg='white',
                font=('Arial',12,'bold')
        ).pack(
                side='left',
                padx=10
        )

        tk.Button(
                banner,
                text='Copy URL',
                command=lambda:self.clipboard_append(
                        self.url.get('1.0','end').strip()
                )
        ).pack(
                side='right',
                padx=4,
                pady=4
        )

        if not self.is_new:
            tk.Button(
                banner,
                text='Delete',
                command=self.delete
            ).pack(
                side='right',
                padx=4,
                pady=4
            )

        tk.Button(
            banner,
                text='Save',
                command=self.save
        ).pack(
                side='right',
                padx=4,
                pady=4
        )

        tk.Button(
                banner,
                text='Cancel',
                command=self.destroy
        ).pack(
                side='right',
                padx=4,
                pady=4
        )
        tk.Label(self,text='Title',bg='#4D148C',fg='white').pack(anchor='w')
        self.title_e=tk.Entry(self)
        self.title_e.pack(fill='x')
        self.title_e.insert(0,link['title'])

        tk.Label(self,text='Collection (read only)',bg='#4D148C',fg='white').pack(anchor='w')
        c=tk.Entry(self)
        c.pack(fill='x')
        c.insert(0,link['collection'])
        c.config(state='readonly')

        self.desc=None
        if show_desc:
            tk.Label(self,text='Description',bg='#4D148C',fg='white').pack(anchor='w')
            self.desc=tk.Text(self,height=3)
            self.desc.pack(fill='x',pady=2)
            self.desc.insert('1.0',description)

        tk.Label(self,text='URL',bg='#4D148C',fg='white').pack(anchor='w')
        self.url=tk.Text(self,height=4,wrap='word')
        self.url.pack(fill='x',pady=2)
        self.url.insert('1.0',link['url'])


    def save(self):
        description=''
        if self.desc:
            description=self.desc.get('1.0','end').strip()
        self.save_cb(self.link,self.title_e.get(),description,self.url.get('1.0','end').strip())
        self.destroy()

    def delete(self):
        if mb.askyesno('Delete Bookmark', f"Delete '{self.link['title']}'?"):
            self.delete_cb(self.link)
            self.destroy()

class CollectionDialog(tk.Toplevel):
    def __init__(self,parent,save_cb,title='',labels=None,heading='Add Collection',delete_cb=None):
        super().__init__(parent)
        self.save_cb=save_cb
        self.delete_cb=delete_cb
        self.title(heading)
        self.transient(parent)
        self.focus_force()
        self.configure(bg='#1e1e1e')
        w,h=800,160
        x=parent.winfo_x()+max(0,(parent.winfo_width()-w)//2)
        y=parent.winfo_y()+max(0,(parent.winfo_height()-h)//2)
        self.geometry(f'{w}x{h}+{x}+{y}')
        banner=tk.Frame(self,bg='#4D148C',height=40)
        banner.pack(fill='x',side='top')
        tk.Label(banner,text=heading,bg='#4D148C',fg='white',font=('Arial',12,'bold')).pack(side='left',padx=10)
        tk.Button(banner,text='Save',command=self.save).pack(side='right',padx=4,pady=4)
        tk.Button(banner,text='Cancel',command=self.destroy).pack(side='right',padx=4,pady=4)
        if delete_cb:
            tk.Button(banner,text='Delete',command=self.delete).pack(side='right',padx=4,pady=4)
        tk.Label(self,text='Title',bg='#4D148C',fg='white').pack(anchor='w')
        self.title_e=tk.Entry(self)
        self.title_e.pack(fill='x')
        self.title_e.insert(0,title)
        tk.Label(self,text='labelIds (comma separated)',bg='#4D148C',fg='white').pack(anchor='w')
        self.labels_e=tk.Entry(self)
        self.labels_e.pack(fill='x')
        self.labels_e.insert(0,', '.join(labels or []))
        self.title_e.focus_set()

    def delete(self):
        if self.delete_cb():
            self.destroy()

    def save(self):
        title=self.title_e.get().strip()
        if not title:
            mb.showerror('Add Collection','Title is required.',parent=self)
            return
        labels=[x.strip() for x in self.labels_e.get().split(',') if x.strip()]
        self.save_cb(title,labels)
        self.destroy()

class App:
    def __init__(self,root):
        self.root=root
        root.title('LinkPilot')
        root.geometry('1500x900')
        root.configure(bg='#4D148C')

        self.state={'favorites':[],'recent':[],'show_urls':False}
        if os.path.exists(STATE_FILE):
            try:
                self.state.update(json.load(open(STATE_FILE)))
            except:
                pass

        self.links=[]
        self.json_file=None
        self.nav_map={}
        self.url_validation={}
        self.validation_generation=0
        self.validation_message=''

        top=tk.Frame(root,bg='#4D148C')
        top.pack(fill='x')

        self.validate_urls=tk.BooleanVar(value=False)
        tk.Checkbutton(top,text='Validate',variable=self.validate_urls,bg='#4D148C',fg='white',
                   command=self.toggle_validation).pack(side='left')

        self.search=tk.StringVar()
        tk.Label(top,text='Search',bg='#4D148C',fg='white').pack(side='left')
        tk.Entry(top,textvariable=self.search,width=50,font=('Arial',12,'bold')).pack(side='left',fill='x',expand=True)
        self.search.trace_add('write',lambda *_: self.refresh_tree())

        self.show_urls=tk.BooleanVar(value=self.state.get('show_urls',False))
        tk.Checkbutton(top,text='Show URLs',variable=self.show_urls,bg='#4D148C',fg='white',
               selectcolor='#251A35',command=self.refresh).pack(side='left')
        tk.Button(top,text='Open Collection',command=self.open_collection).pack(side='left')

        self.stat=tk.Label(top,bg='#4D148C',fg='#FFFFFF')
        self.stat.pack(side='right')

        pan=tk.PanedWindow(root,orient='horizontal')
        pan.pack(fill='both',expand=True)

        left=tk.Frame(pan,bg='#4D148C')
        pan.add(left,width=350)

        style = ttk.Style()
        style.theme_use('clam')

        style.configure(
            'Treeview',
            background='#FFFFFF',
            fieldbackground='#FFFFFF',
            foreground='#000000',
            rowheight=30,
            font=('Arial', 13, 'bold')
        )

        style.configure(
            'Treeview.Heading',
            background='#4D148C',
            foreground='#FFFFFF',
            font=('Arial', 13, 'bold')
        )

        style.map(
            'Treeview',
            background=[('selected', '#4D148C')],
            foreground=[('selected', '#FFFFFF')]
        )

        self.collections=ttk.Treeview(left,show='tree')
        ysb=ttk.Scrollbar(left,orient='vertical',command=self.collections.yview)
        self.collections.configure(yscrollcommand=ysb.set)
        hint=tk.Label(left,text='Right-click a collection to add a link; click Collections to add a collection, click a collection name to edit it',bg='#4D148C',fg='white',font=('Arial',10,'italic'),justify='left',anchor='w',wraplength=330)
        hint.pack(side='bottom',fill='x')
        hint.bind('<Configure>',lambda e:hint.config(wraplength=max(50,e.width-10)))
        self.collections.pack(side='left',fill='both',expand=True)
        ysb.pack(side='right',fill='y')
        self.collections.tag_configure('invalid',foreground='#D00000')
        self.collections.bind('<<TreeviewSelect>>',self.collection_selected)
        self.collections.bind('<Button-3>',self.add_bookmark)
        self.collections.bind('<ButtonRelease-1>',self.add_collection_click)

        rf=tk.Frame(pan)
        pan.add(rf)

        self.tree=ttk.Treeview(rf,columns=('Title','Collection','URL'),show='tree headings')
        for c,w in [('#0',50),('Title',500),('Collection',280),('URL',500)]:
            self.tree.heading(c,text='Fav' if c=='#0' else c)
            self.tree.column(c,width=w,anchor='center' if c=='#0' else 'w')
        self.star_icons={
            (False,False):make_star_icon(root,'#777777',False),
            (True,False):make_star_icon(root,'#FFB000',True),
            (False,True):make_star_icon(root,'#D00000',True),
            (True,True):make_star_icon(root,'#D00000',True)
        }

        sb=ttk.Scrollbar(rf,orient='vertical',command=self.tree.yview)
        xsb=ttk.Scrollbar(rf,orient='horizontal',command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb.set,xscrollcommand=xsb.set)
        sb.pack(side='right',fill='y')
        xsb.pack(side='bottom',fill='x')
        self.tree.pack(side='left',fill='both',expand=True)
        rownum = 0
        self.tree.tag_configure(
            'odd',
            background='#FFFFFF'
        )

        self.tree.tag_configure(
            'even',
            background='#E8E0F3'
        )

        self.tree.bind('<ButtonRelease-1>',self.click)
        self.tree.bind('<Button-3>',self.edit)

        self.load_newest()

    def save_state(self):
        self.state['show_urls']=self.show_urls.get()
        json.dump(self.state,open(STATE_FILE,'w'),indent=2)

    def load_newest(self):
        files=glob.glob(os.path.join(APP_DIR,'link-export*.json'))
        if not files:
            self.json_file=os.path.join(APP_DIR,'link-export.json')
            self.data={'groups':[{'title':'Default','lists':[]}]}
            with open(self.json_file,'w',encoding='utf-8') as file:
                json.dump(self.data,file,indent=2)
        else:
            self.json_file=max(files,key=os.path.getmtime)
        self.load_json()

    def load_json(self):
        self.data=json.load(open(self.json_file,encoding='utf-8'))
        self.links=[]
        counts={}
        bid=1

        for g in self.data.get('groups',[]):
            for lst in g.get('lists',[]):
                col=lst.get('title','')
                counts[col]=counts.get(col,0)+len(lst.get('cards',[]))
                for card in lst.get('cards',[]):
                    self.links.append({'id':bid,'card':card,'title':card.get('customTitle') or card.get('title',''),'description':(card.get('customDescription') or '')+' '+(card.get('description') or ''),'url':card.get('url',''),'collection':col})
                    bid+=1

        self.collection_counts=counts
        self.build_tree(counts)
        self.refresh()

    def build_tree(self,counts):
        self.collections.delete(*self.collections.get_children())
        self.nav_map={}
        term=self.search.get().lower().strip()
        fav=self.collections.insert('', 'end', text='⭐ Favorites', open=True)
        recent=self.collections.insert('', 'end', text='🕒 Recent', open=True)
        allnode=self.collections.insert('', 'end', text='All Links', open=True)
        colroot=self.collections.insert('', 'end', text='Collections', open=True)
        self.nav_map[fav]=('favorites',None)
        self.nav_map[recent]=('recent',None)
        self.nav_map[allnode]=('all',None)
        self.nav_map[colroot]=('collections_root',None)
        for col,count in sorted(counts.items()):
            collection_matches=term and term in col.lower()
            matching=[]
            collection_links=[link for link in self.links if link['collection']==col]
            for link in collection_links:
                st=(link['title']+' '+link['description']+' '+link['url']).lower()
                if not term or collection_matches or term in st:
                    matching.append(link)
            if term and not matching:
                continue
            has_invalid=any(self.url_validation.get((col,link['url'].strip())) is False
                            for link in collection_links)
            parent=self.collections.insert(colroot,'end',text=f'{col} ({count})',
                                           open=bool(term),
                                           tags=('invalid',) if has_invalid else ())
            self.nav_map[parent]=('collection',col)
            for link in sorted(matching,key=lambda x:x['title'].lower()):
                child=self.collections.insert(parent,'end',text=link['title'])
                self.nav_map[child]=('link',link['url'])
        self.collections.selection_set(allnode)

    def refresh_tree(self):
        if hasattr(self, "collection_counts"):
            self.build_tree(self.collection_counts)
        self.refresh()

    def collection_selected(self,event=None):
        self.refresh()
        if self.validate_urls.get():
            self.validation_generation+=1
            if not self.validate_selected_collection():
                self.validation_message='Select one collection to validate'
                self.refresh()

    def toggle_validation(self):
        if self.validate_urls.get():
            dialog=TokenDialog(self.root,getattr(self,'access_token',''))
            self.root.wait_window(dialog)
            if dialog.result is None:
                self.validate_urls.set(False)
                self.toggle_validation()
                return
            self.access_token=dialog.result
            if self.access_token:
                print(describe_token(self.access_token), file=sys.stderr)
            if not self.validate_selected_collection():
                self.validation_message='Select one collection to validate'
                self.refresh()
            return
        self.validation_generation+=1
        self.url_validation.clear()
        self.validation_message=''
        for item,(kind,value) in self.nav_map.items():
            if kind=='collection':
                self.collections.item(item,tags=())
        self.refresh()

    def validate_selected_collection(self):
        if not self.validate_urls.get():
            return False
        selection=self.collections.selection()
        if not selection:
            return False
        kind,collection=self.nav_map.get(selection[0],('all',None))
        if kind!='collection':
            return False

        token=getattr(self,'access_token','')

        urls=list(dict.fromkeys(link['url'].strip() for link in self.links
                                if link['collection']==collection))
        self.validation_generation+=1
        generation=self.validation_generation
        for url in urls:
            self.url_validation.pop((collection,url),None)
        self.validation_message=f'Checking {collection}: 0/{len(urls)}'
        self.update_collection_label(collection)
        self.refresh()

        def worker():
            results={}
            for completed,url in enumerate(urls,1):
                if generation!=self.validation_generation:
                    break
                results[url]=is_url_valid(url,token)
                self.root.after(0,lambda count=completed:self.update_validation_progress(
                    collection,count,len(urls),generation))
            self.root.after(0,lambda:self.finish_validation(collection,results,generation))

        threading.Thread(target=worker,daemon=True).start()
        return True

    def update_validation_progress(self,collection,completed,total,generation):
        if generation==self.validation_generation and self.validate_urls.get():
            self.validation_message=f'Checking {collection}: {completed}/{total}'
            self.stat.config(text=self.validation_message)

    def finish_validation(self,collection,results,generation):
        if generation!=self.validation_generation or not self.validate_urls.get():
            return
        for url,is_valid in results.items():
            self.url_validation[(collection,url)]=is_valid
        invalid_count=sum(not is_valid for is_valid in results.values())
        self.validation_message=f'{collection}: {invalid_count} invalid URL(s)'
        self.update_collection_label(collection)
        self.refresh()

    def update_collection_label(self,collection):
        has_invalid=any(self.url_validation.get((collection,link['url'].strip())) is False
                        for link in self.links if link['collection']==collection)
        for item,(kind,value) in self.nav_map.items():
            if kind=='collection' and value==collection:
                self.collections.item(item,tags=('invalid',) if has_invalid else ())

    def find(self,item):
        vals=self.tree.item(item,'values')
        if not vals or len(vals)<3:
            return None
        url=vals[2]
        for l in self.links:
            if l['url']==url:
                return l
        return None

    def refresh(self):
        self.save_state()
        self.tree['displaycolumns']=('Title','Collection','URL') if self.show_urls.get() else ('Title','Collection')
        for i in self.tree.get_children():
            self.tree.delete(i)

        rows=self.links
        sel=self.collections.selection()
        if sel:
            kind,val=self.nav_map.get(sel[0],('all',None))
            if kind=='favorites':
                rows=[r for r in rows if r['url'] in self.state['favorites']]
            elif kind=='recent':
                rows=[r for r in rows if r['url'] in self.state['recent']]
            elif kind=='collection':
                rows=[r for r in rows if r['collection']==val]
            elif kind=='link':
                rows=[r for r in rows if r['url']==val]
        rownum = 0
        term=self.search.get().lower()
        for r in rows:
            if term and term not in (r['title']+' '+r['description']+' '+r['url']+' '+r['collection']).lower():
                continue
            tag = 'odd' if rownum % 2 == 0 else 'even'
            favorite=r['url'] in self.state['favorites']
            invalid=self.url_validation.get((r['collection'],r['url'].strip())) is False
            self.tree.insert('','end',text='',image=self.star_icons[(favorite,invalid)],
                values=(r['title'],r['collection'],r['url']),
                tags=(tag,)
            )
            rownum += 1
        self.stat.config(text=self.validation_message or
                 f"{len(self.links)} Links | {len(self.state['favorites'])} Favorites")

    def click(self,e):
        item=self.tree.identify_row(e.y)
        col=self.tree.identify_column(e.x)
        if not item:
            return
        l=self.find(item)
        if not l:
            return
        if col=='#0':
            f=self.state['favorites']
            u=l['url']
            f.remove(u) if u in f else f.append(u)
            self.save_state()
            self.refresh()
        elif col=='#1' and l.get('url'):
            webbrowser.open_new_tab(l['url'])
            self.state['recent']=[l['url']]+[x for x in self.state['recent'] if x!=l['url']][:49]
            self.save_state()

    def edit(self,e):
        item=self.tree.identify_row(e.y)
        if not item:
            return
        l=self.find(item)
        if l:
            EditDialog(self.root,l,self.save_link,self.delete_link)

    def add_bookmark(self,e):
        item=self.collections.identify_row(e.y)
        if not item:
            return
        kind,val=self.nav_map.get(item,('all',None))
        if kind!='collection':
            return
        new={'new':True,'card':None,'title':'','description':'','url':'','collection':val}
        EditDialog(self.root,new,self.save_link,self.delete_link)

    def add_collection_click(self,e):
        if 'indicator' in self.collections.identify_element(e.x,e.y):
            return
        item=self.collections.identify_row(e.y)
        if not item:
            return
        kind,val=self.nav_map.get(item,('all',None))
        if kind=='collections_root':
            CollectionDialog(self.root,self.save_collection)
        elif kind=='collection':
            lst=self.find_list(val)
            if lst is not None:
                CollectionDialog(self.root,lambda t,l:self.update_collection(lst,t,l),lst.get('title',''),lst.get('labelIds',[]),'Edit Collection',lambda:self.delete_collection(lst))

    def find_list(self,title):
        for group in self.data.get('groups',[]):
            for lst in group.get('lists',[]):
                if lst.get('title','')==title:
                    return lst
        return None

    def delete_collection(self,lst):
        cards=lst.get('cards',[])
        if not mb.askyesno('Delete Collection',f"Delete '{lst.get('title','')}' and its {len(cards)} links? This cannot be undone."):
            return False
        bk=os.path.join(BACKUP_DIR,os.path.basename(self.json_file).replace('.json','_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.json'))
        shutil.copy2(self.json_file,bk)
        urls={c.get('url','') for c in cards}
        for group in self.data.get('groups',[]):
            if lst in group.get('lists',[]):
                group['lists']=[l for l in group['lists'] if l is not lst]
                break
        with open(self.json_file,'w',encoding='utf-8') as file:
            json.dump(self.data,file,indent=2)
        remaining={l['url'] for l in self.links if l['card'] not in cards}
        for key in ('favorites','recent'):
            self.state[key]=[u for u in self.state[key] if u not in urls or u in remaining]
        self.load_json()
        return True

    def update_collection(self,lst,title,labels):
        bk=os.path.join(BACKUP_DIR,os.path.basename(self.json_file).replace('.json','_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.json'))
        shutil.copy2(self.json_file,bk)
        lst['title']=title
        lst['labelIds']=labels
        with open(self.json_file,'w',encoding='utf-8') as file:
            json.dump(self.data,file,indent=2)
        self.load_json()

    def save_collection(self,title,labels):
        groups=self.data.get('groups',[])
        if not groups:
            mb.showerror('Add Collection','No group found in the loaded export.')
            return
        bk=os.path.join(BACKUP_DIR,os.path.basename(self.json_file).replace('.json','_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.json'))
        shutil.copy2(self.json_file,bk)
        groups[0].setdefault('lists',[]).append({'title':title,'labelIds':labels,'cards':[]})
        with open(self.json_file,'w',encoding='utf-8') as file:
            json.dump(self.data,file,indent=2)
        self.load_json()

    def save_link(self,link,title,desc,url):
        bk=os.path.join(BACKUP_DIR,os.path.basename(self.json_file).replace('.json','_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.json'))
        shutil.copy2(self.json_file,bk)
        if link.get('new'):
            card={'title':title,'url':url,'customTitle':title,'customDescription':desc}
            for group in self.data.get('groups',[]):
                for listing in group.get('lists',[]):
                    if listing.get('title','')==link['collection']:
                        listing.setdefault('cards',[]).append(card)
                        with open(self.json_file,'w',encoding='utf-8') as file:
                            json.dump(self.data,file,indent=2)
                        self.load_json()
                        return
            mb.showerror('Add Bookmark','Collection not found in the loaded export.')
            return
        link['card']['customTitle']=title
        link['card']['customDescription']=desc
        link['card']['url']=url
        link['title']=title
        link['description']=desc
        link['url']=url
        json.dump(self.data,open(self.json_file,'w',encoding='utf-8'),indent=2)
        self.load_json()

    def delete_link(self,link):
        target_cards=None
        target_index=None
        for group in self.data.get('groups',[]):
            for listing in group.get('lists',[]):
                cards=listing.get('cards',[])
                for index,card in enumerate(cards):
                    if card is link['card']:
                        target_cards=cards
                        target_index=index
                        break
                if target_cards is not None:
                    break
            if target_cards is not None:
                break

        if target_cards is None:
            mb.showerror('Delete Bookmark','The bookmark could not be found in the loaded export.')
            return

        bk=os.path.join(BACKUP_DIR,os.path.basename(self.json_file).replace('.json','_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.json'))
        shutil.copy2(self.json_file,bk)
        target_cards.pop(target_index)
        with open(self.json_file,'w',encoding='utf-8') as file:
            json.dump(self.data,file,indent=2)
        self.load_json()

    def open_collection(self):
        sel=self.collections.selection()
        if not sel:
            return
        kind,val=self.nav_map.get(sel[0],('all',None))
        if kind!='collection':
            return
        for l in self.links:
            if l['collection']==val and l['url']:
                webbrowser.open_new_tab(l['url'])

root=tk.Tk()
App(root)
root.mainloop()
