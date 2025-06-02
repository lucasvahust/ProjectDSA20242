import tkinter as tk
from tkinter import filedialog, ttk, messagebox, Toplevel
import pandas as pd
from datetime import datetime
import os
import re
import uuid
import logging

# Thiết lập logging
logging.basicConfig(filename='library.log', level=logging.INFO)

# ========================= Lớp Node và LinkedList (tái sử dụng) =========================
class Node:
    def __init__(self, data):
        self.data = data
        self.next = None

class LinkedList:
    def __init__(self):
        self.head = None

    def appendLast(self, data):
        new_node = Node(data)
        if not self.head:
            self.head = new_node
            return
        current = self.head
        while current.next:
            current = current.next
        current.next = new_node

    def removeFirst(self):
        if self.head:
            self.head = self.head.next

    def removeLast(self):
        if not self.head:
            return
        if not self.head.next:
            self.head = None
            return
        current = self.head
        while current.next.next:
            current = current.next
        current.next = None

    def get_list(self):
        result = []
        current = self.head
        while current:
            result.append(current.data)
            current = current.next
        return result

# ========================= Lớp Teacher và Student =========================
class Teacher:
    def __init__(self, username, password, name):
        self.username = username
        self.password = password
        self.name = name

class Student:
    def __init__(self, username, password, name):
        self.username = username
        self.password = password
        self.name = name

# ========================= Lớp TeacherList và StudentList =========================
class TeacherList:
    def __init__(self, file_path):
        self.file_path = file_path
        self.teachers = LinkedList()
        self.load_data()

    def load_data(self):
        try:
            if os.path.exists(self.file_path):
                df = pd.read_excel(self.file_path)
                if df.empty:
                    self.save_to_excel()
                    return
                if not all(col in df.columns for col in ["Username", "Password", "Name"]):
                    raise ValueError("File Excel không đúng định dạng")
                for _, row in df.iterrows():
                    teacher = Teacher(row["Username"], row["Password"], row["Name"])
                    self.teachers.appendLast(teacher)
            else:
                self.save_to_excel()
        except Exception as e:
            logging.error(f"Không thể đọc file {self.file_path}: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể đọc file: {self.file_path}\n{str(e)}")

    def save_to_excel(self):
        try:
            df = pd.DataFrame([(t.username, t.password, t.name)
                               for t in self.teachers.get_list()],
                              columns=["Username", "Password", "Name"])
            df.to_excel(self.file_path, index=False)
        except Exception as e:
            logging.error(f"Lỗi khi lưu file {self.file_path}: {str(e)}")
            messagebox.showerror("Lỗi", f"Lỗi khi lưu file: {str(e)}")

    def authenticate(self, username, password):
        current = self.teachers.head
        while current:
            if current.data.username == username and current.data.password == password:
                return True
            current = current.next
        return False

class StudentList:
    def __init__(self, file_path):
        self.file_path = file_path
        self.students = LinkedList()
        self.load_data()

    def load_data(self):
        try:
            if os.path.exists(self.file_path):
                df = pd.read_excel(self.file_path)
                if df.empty:
                    self.save_to_excel()
                    return
                if not all(col in df.columns for col in ["Username", "Password", "Name"]):
                    raise ValueError("File Excel không đúng định dạng")
                for _, row in df.iterrows():
                    student = Student(row["Username"], row["Password"], row["Name"])
                    self.students.appendLast(student)
            else:
                self.save_to_excel()
        except Exception as e:
            logging.error(f"Không thể đọc file {self.file_path}: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể đọc file: {self.file_path}\n{str(e)}")

    def save_to_excel(self):
        try:
            df = pd.DataFrame([(s.username, s.password, s.name)
                               for s in self.students.get_list()],
                              columns=["Username", "Password", "Name"])
            df.to_excel(self.file_path, index=False)
        except Exception as e:
            logging.error(f"Lỗi khi lưu file {self.file_path}: {str(e)}")
            messagebox.showerror("Lỗi", f"Lỗi khi lưu file: {str(e)}")

    def authenticate(self, username, password):
        current = self.students.head
        while current:
            if current.data.username == username and current.data.password == password:
                return current.data.name  # Trả về tên của sinh viên
            current = current.next
        return False

    def get_student_name(self, username):
        current = self.students.head
        while current:
            if current.data.username == username:
                return current.data.name
            current = current.next
        return None

# ========================= Lớp Book =========================
class Book:
    def __init__(self, Ma_sach, ten_sach, tac_gia, nha_xuat_ban, nam_xuat_ban, so_ISBN):
        self.Ma_sach = Ma_sach
        self.ten_sach = ten_sach
        self.tac_gia = tac_gia
        self.nha_xuat_ban = nha_xuat_ban
        self.nam_xuat_ban = nam_xuat_ban
        self.so_ISBN = so_ISBN

# ========================= Lớp BookList =========================
class BookList:
    def __init__(self, file_path):
        self.file_path = file_path
        self.books = LinkedList()
        self.load_data()
        self.sort_by_ma_sach()

    def load_data(self):
        try:
            if os.path.exists(self.file_path):
                df = pd.read_excel(self.file_path)
                if df.empty:
                    self.save_to_excel()
                    return
                if not all(col in df.columns for col in ["Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Năm xuất bản", "Số ISBN"]):
                    raise ValueError("File Excel không đúng định dạng")
                for _, row in df.iterrows():
                    book = Book(row["Mã sách"], row["Tên sách"], row["Tác giả"], row["Nhà xuất bản"], row["Năm xuất bản"], row["Số ISBN"])
                    self.books.appendLast(book)
                self.sort_by_ma_sach()
            else:
                self.save_to_excel()
        except Exception as e:
            logging.error(f"Không thể đọc file {self.file_path}: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể đọc file: {self.file_path}\n{str(e)}")

    def save_to_excel(self):
        try:
            df = pd.DataFrame([(b.Ma_sach, b.ten_sach, b.tac_gia, b.nha_xuat_ban, b.nam_xuat_ban, b.so_ISBN)
                               for b in self.books.get_list()],
                              columns=["Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Năm xuất bản", "Số ISBN"])
            df.to_excel(self.file_path, index=False)
        except PermissionError:
            logging.error(f"Không thể lưu file {self.file_path}: Permission denied")
            messagebox.showerror("Lỗi", "Không thể lưu file! Vui lòng đóng file Excel nếu đang mở hoặc kiểm tra quyền ghi.")
        except Exception as e:
            logging.error(f"Lỗi khi lưu file {self.file_path}: {str(e)}")
            messagebox.showerror("Lỗi", f"Lỗi khi lưu file: {str(e)}")

    def extract_number(self, ma):
        so = ''.join(ch for ch in str(ma) if ch.isdigit())
        return int(so) if so else 0

    def sort_by_ma_sach(self):
        if self.books.head is None or self.books.head.next is None:
            return
        swapped = True
        while swapped:
            swapped = False
            prev = None
            curr = self.books.head
            next_node = curr.next
            while next_node:
                if self.extract_number(curr.data.Ma_sach) > self.extract_number(next_node.data.Ma_sach):
                    if prev is None:
                        self.books.head = next_node
                    else:
                        prev.next = next_node
                    curr.next = next_node.next
                    next_node.next = curr
                    swapped = True
                    prev = next_node
                    next_node = curr.next
                else:
                    prev = curr
                    curr = next_node
                    next_node = next_node.next

    def search_book_multi(self):
        search_term = entry_search.get().strip()
        if not search_term:
            messagebox.showerror("Lỗi", "Vui lòng nhập từ khóa tìm kiếm!")
            return
        results = LinkedList()
        current = self.books.head
        while current:
            book = current.data
            if (search_term.lower() in str(book.Ma_sach).lower() or
                search_term.lower() in str(book.ten_sach).lower() or
                search_term.lower() in str(book.tac_gia).lower() or
                search_term.lower() in str(book.nha_xuat_ban).lower() or
                search_term.lower() in str(book.nam_xuat_ban).lower() or
                search_term.lower() in str(book.so_ISBN).lower()):
                results.appendLast(book)
            current = current.next
        if not results.head:
            messagebox.showwarning("Không tìm thấy", f"Không tìm thấy sách khớp với: {search_term}")
            update_treeview(self.books)
            entry_search.delete(0, tk.END)
            return
        update_treeview(results)
        entry_search.delete(0, tk.END)

    def add_book(self):
        new_ma_sach = entry_code.get().strip()
        new_ten_sach = entry_title.get().strip()
        new_tac_gia = entry_author.get().strip()
        new_nha_xuat_ban = entry_publisher.get().strip()
        new_nam_xuat_ban = entry_year.get().strip()
        new_so_isbn = entry_isbn.get().strip()

        if not all([new_ma_sach, new_ten_sach, new_tac_gia, new_nha_xuat_ban, new_nam_xuat_ban, new_so_isbn]):
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin sách!")
            return
        if len(new_ma_sach) > 50 or len(new_ten_sach) > 100 or len(new_tac_gia) > 50 or len(new_nha_xuat_ban) > 50:
            messagebox.showerror("Lỗi", "Thông tin nhập vào quá dài!")
            return
        if not re.match(r"^[a-zA-Z0-9\s]+$", new_ma_sach):
            messagebox.showerror("Lỗi", "Mã sách chỉ được chứa chữ cái, số và khoảng trắng!")
            return
        if not new_nam_xuat_ban.isdigit() or int(new_nam_xuat_ban) < 0:
            messagebox.showerror("Lỗi", "Năm xuất bản phải là số dương!")
            return
        if not re.match(r"^\d{10}|\d{13}$", new_so_isbn.replace("-", "")):
            messagebox.showerror("Lỗi", "Số ISBN phải có 10 hoặc 13 chữ số!")
            return

        current = self.books.head
        while current:
            if current.data.Ma_sach == new_ma_sach:
                messagebox.showerror("Lỗi", f"Mã sách {new_ma_sach} đã tồn tại!")
                return
            current = current.next

        new_book = Book(new_ma_sach, new_ten_sach, new_tac_gia, new_nha_xuat_ban, new_nam_xuat_ban, new_so_isbn)
        self.books.appendLast(new_book)
        self.sort_by_ma_sach()
        self.save_to_excel()
        update_treeview(self.books)
        clear_entries()

    def edit_book(self):
        edit_window = Toplevel(root)
        edit_window.title("Sửa biên mục sách")
        tk.Label(edit_window, text="Nhập mã sách để sửa:").pack()
        entry_ma_sach = tk.Entry(edit_window)
        entry_ma_sach.pack()

        def confirm_edit():
            ma_sach = entry_ma_sach.get().strip()
            current = self.books.head
            found = False
            while current:
                book = current.data
                if book.Ma_sach == ma_sach:
                    found = True
                    edit_detail_window = Toplevel(edit_window)
                    edit_detail_window.title("Sửa thông tin sách")
                    labels = ["Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Năm xuất bản", "Số ISBN"]
                    entries = []
                    for i, text in enumerate(labels):
                        tk.Label(edit_detail_window, text=text, width=20, anchor='w').grid(row=i, column=0, padx=10, pady=5)
                        entry = tk.Entry(edit_detail_window, width=30)
                        entry.grid(row=i, column=1, padx=10, pady=5)
                        entries.append(entry)
                    entries[0].insert(0, book.Ma_sach)
                    entries[1].insert(0, book.ten_sach)
                    entries[2].insert(0, book.tac_gia)
                    entries[3].insert(0, book.nha_xuat_ban)
                    entries[4].insert(0, book.nam_xuat_ban)
                    entries[5].insert(0, book.so_ISBN)

                    def save_edit():
                        new_ma_sach, new_ten_sach, new_tac_gia, new_nha_xuat_ban, new_nam_xuat_ban, new_so_isbn = [e.get().strip() for e in entries]
                        if not all([new_ma_sach, new_ten_sach, new_tac_gia, new_nha_xuat_ban, new_nam_xuat_ban, new_so_isbn]):
                            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin sách!")
                            return
                        if len(new_ma_sach) > 50 or len(new_ten_sach) > 100 or len(new_tac_gia) > 50 or len(new_nha_xuat_ban) > 50:
                            messagebox.showerror("Lỗi", "Thông tin nhập vào quá dài!")
                            return
                        if not re.match(r"^[a-zA-Z0-9\s]+$", new_ma_sach):
                            messagebox.showerror("Lỗi", "Mã sách chỉ được chứa chữ cái, số và khoảng trắng!")
                            return
                        if not new_nam_xuat_ban.isdigit() or int(new_nam_xuat_ban) < 0:
                            messagebox.showerror("Lỗi", "Năm xuất bản phải là số dương!")
                            return
                        if not re.match(r"^\d{10}|\d{13}$", new_so_isbn.replace("-", "")):
                            messagebox.showerror("Lỗi", "Số ISBN phải có 10 hoặc 13 chữ số!")
                            return
                        temp = self.books.head
                        while temp:
                            if temp.data.Ma_sach == new_ma_sach and temp.data.Ma_sach != book.Ma_sach:
                                messagebox.showerror("Lỗi", f"Mã sách {new_ma_sach} đã tồn tại!")
                                return
                            temp = temp.next
                        book.Ma_sach = new_ma_sach
                        book.ten_sach = new_ten_sach
                        book.tac_gia = new_tac_gia
                        book.nha_xuat_ban = new_nha_xuat_ban
                        book.nam_xuat_ban = new_nam_xuat_ban
                        book.so_ISBN = new_so_isbn
                        self.sort_by_ma_sach()
                        self.save_to_excel()
                        update_treeview(self.books)
                        edit_detail_window.destroy()
                        edit_window.destroy()

                    tk.Button(edit_detail_window, text="Lưu thay đổi", command=save_edit).grid(row=6, column=1, pady=10)
                    break
                current = current.next
            if not found:
                messagebox.showwarning("Không tìm thấy", f"Không tìm thấy sách với mã: {ma_sach}")
                entry_ma_sach.delete(0, tk.END)

        tk.Button(edit_window, text="Xác nhận", command=confirm_edit).pack(pady=10)

    def delete_book(self):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn sách để xóa.")
            return
        values = tree.item(selected[0])['values']
        ma_sach_selected = values[0]
        if not messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa sách {ma_sach_selected}?"):
            return
        prev = None
        current = self.books.head
        while current:
            if current.data.Ma_sach == ma_sach_selected:
                if prev is None:
                    self.books.removeFirst()
                elif current.next is None:
                    self.books.removeLast()
                else:
                    prev.next = current.next
                break
            prev = current
            current = current.next
        self.save_to_excel()
        update_treeview(self.books)

# ========================= Lớp BookDetail & Danh sách =========================
class BookDetail:
    def __init__(self, so_nhap_kho, ma_sach, tinh_trang):
        self.so_nhap_kho = so_nhap_kho
        self.ma_sach = ma_sach
        self.tinh_trang = tinh_trang

class BookDetailList:
    def __init__(self, file_path):
        self.file_path = file_path
        self.details = LinkedList()
        self.load_data()
        self.sort_by_so_nhap_kho()
        self.save_detail_to_excel()

    def load_data(self):
        try:
            if os.path.exists(self.file_path):
                df = pd.read_excel(self.file_path)
                if df.empty:
                    self.save_detail_to_excel()
                    return
                if not all(col in df.columns for col in ["Số nhập kho", "Mã sách", "Trạng thái"]):
                    raise ValueError("File Excel không đúng định dạng")
                for _, row in df.iterrows():
                    if row["Số nhập kho"] and row["Mã sách"] and row["Trạng thái"]:
                        self.details.appendLast(BookDetail(row["Số nhập kho"], row["Mã sách"], row["Trạng thái"]))
                self.sort_by_so_nhap_kho()
            else:
                self.save_detail_to_excel()
        except Exception as e:
            logging.error(f"Không thể đọc file {self.file_path}: {str(e)}")
            messagebox.showerror("Lỗi", f"Không thể đọc file: {self.file_path}\n{str(e)}")

    def save_detail_to_excel(self):
        try:
            df = pd.DataFrame([(d.so_nhap_kho, d.ma_sach, d.tinh_trang) for d in self.details.get_list()],
                              columns=["Số nhập kho", "Mã sách", "Trạng thái"])
            df.to_excel(self.file_path, index=False)
        except PermissionError:
            logging.error(f"Không thể lưu file {self.file_path}: Permission denied")
            messagebox.showerror("Lỗi", "Không thể lưu file! Vui lòng đóng file Excel nếu đang mở hoặc kiểm tra quyền ghi.")
        except Exception as e:
            logging.error(f"Lỗi khi lưu file {self.file_path}: {str(e)}")
            messagebox.showerror("Lỗi", f"Lỗi khi lưu file: {str(e)}")

    def sort_by_so_nhap_kho(self):
        if self.details.head is None or self.details.head.next is None:
            return
        swapped = True
        while swapped:
            swapped = False
            prev = None
            curr = self.details.head
            next_node = curr.next
            while next_node:
                if curr.data.so_nhap_kho > next_node.data.so_nhap_kho:
                    if prev is None:
                        self.details.head = next_node
                    else:
                        prev.next = next_node
                    curr.next = next_node.next
                    next_node.next = curr
                    swapped = True
                    prev = next_node
                    next_node = curr.next
                else:
                    prev = curr
                    curr = next_node
                    next_node = next_node.next

    def add_detail(self, detail):
        if not re.match(r"^[a-zA-Z0-9\s]+$", detail.so_nhap_kho):
            messagebox.showerror("Lỗi", "Số nhập kho chỉ được chứa chữ cái, số và khoảng trắng!")
            return
        if not re.match(r"^[a-zA-Z0-9\s]+$", detail.ma_sach):
            messagebox.showerror("Lỗi", "Mã sách chỉ được chứa chữ cái, số và khoảng trắng!")
            return
        current = book_list.books.head
        found = False
        while current:
            if current.data.Ma_sach == detail.ma_sach:
                found = True
                break
            current = current.next
        if not found:
            messagebox.showerror("Lỗi", f"Mã sách {detail.ma_sach} không tồn tại!")
            return
        current = self.details.head
        while current:
            if current.data.so_nhap_kho == detail.so_nhap_kho:
                messagebox.showerror("Lỗi", f"Số nhập kho {detail.so_nhap_kho} đã tồn tại!")
                return
            current = current.next

        self.details.appendLast(detail)
        self.sort_by_so_nhap_kho()
        self.save_detail_to_excel()
        self.update_detail_treeview()

    def delete_detail(self, so_nhap_kho):
        if not messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa sách với số nhập kho {so_nhap_kho}?"):
            return
        prev = None
        curr = self.details.head
        while curr:
            if curr.data.so_nhap_kho == so_nhap_kho:
                if prev is None:
                    self.details.removeFirst()
                elif curr.next is None:
                    self.details.removeLast()
                else:
                    prev.next = curr.next
                self.sort_by_so_nhap_kho()
                self.save_detail_to_excel()
                self.update_detail_treeview()
                return
            prev = curr
            curr = curr.next
        messagebox.showwarning("Không tìm thấy", f"Không tìm thấy sách với số nhập kho: {so_nhap_kho}")

    def update_status(self, so_nhap_kho, new_status):
        if new_status not in ["Có sẵn", "Đã mượn"]:
            messagebox.showerror("Lỗi", "Trạng thái phải là 'Có sẵn' hoặc 'Đã mượn'!")
            return
        current = self.details.head
        while current:
            if current.data.so_nhap_kho == so_nhap_kho:
                current.data.tinh_trang = new_status
                self.sort_by_so_nhap_kho()
                self.save_detail_to_excel()
                self.update_detail_treeview()
                return
            current = current.next
        messagebox.showwarning("Không tìm thấy", f"Không tìm thấy sách với số nhập kho: {so_nhap_kho}")

    def update_detail_treeview(self, filter_ma_sach=None):
        if tree_detail is None:  # Nếu tree_detail chưa được khởi tạo, bỏ qua
            return
        for row in tree_detail.get_children():
            tree_detail.delete(row)
        current = self.details.head
        while current:
            d = current.data
            if filter_ma_sach is None or d.ma_sach == filter_ma_sach:
                tree_detail.insert("", "end", values=(d.so_nhap_kho, d.ma_sach, d.tinh_trang))
            current = current.next
# ========================= Hàm hỗ trợ =========================
def clear_entries():
    entry_code.delete(0, tk.END)
    entry_title.delete(0, tk.END)
    entry_author.delete(0, tk.END)
    entry_publisher.delete(0, tk.END)
    entry_year.delete(0, tk.END)
    entry_isbn.delete(0, tk.END)

def update_treeview(linked_list):
    for row in tree.get_children():
        tree.delete(row)
    current = linked_list.head
    if not current:
        messagebox.showinfo("Thông báo", "Danh sách sách trống.")
        return
    while current:
        book = current.data
        tree.insert("", "end", values=(book.Ma_sach, book.ten_sach, book.tac_gia,
                                       book.nha_xuat_ban, book.nam_xuat_ban, book.so_ISBN))
        current = current.next

def approve_requests():
    if not os.path.exists("borrow_requests.xlsx"):
        messagebox.showinfo("Thông báo", "Không có yêu cầu nào.")
        return
    window = tk.Toplevel(root)
    window.title("Duyệt yêu cầu mượn sách")
    window.geometry("600x400")

    tree = ttk.Treeview(window, columns=["ID", "Số nhập kho", "Mã sách", "Người mượn", "Thời gian", "Trạng thái"], show="headings")
    for col in ["ID", "Số nhập kho", "Mã sách", "Người mượn", "Thời gian", "Trạng thái"]:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    vsb = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=vsb.set)

    # Load the DataFrame and store it in a nonlocal variable
    try:
        df = pd.read_excel("borrow_requests.xlsx")
        for _, row in df.iterrows():
            tree.insert("", "end", values=(row["ID"], row["Số nhập kho"], row["Mã sách"], row["Người mượn"], row["Thời gian"], row["Trạng thái"]))
    except Exception as e:
        logging.error(f"Lỗi khi đọc borrow_requests.xlsx: {str(e)}")
        messagebox.showerror("Lỗi", f"Lỗi khi đọc yêu cầu: {str(e)}")
        return

    def approve():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn một yêu cầu để duyệt.")
            return
        values = tree.item(selected[0])['values']
        request_id = values[0]
        snk = values[1]
        status = values[5]
        if status != "Chờ duyệt":
            messagebox.showwarning("Lỗi", f"Yêu cầu này đã được xử lý (Trạng thái: {status}).")
            return
        try:
            # Update the DataFrame
            df.loc[df["ID"] == request_id, "Trạng thái"] = "Đã duyệt"
            detail_list.update_status(snk, "Đã mượn")
            # Save the updated DataFrame to Excel
            df.to_excel("borrow_requests.xlsx", index=False)
            messagebox.showinfo("Thành công", "Yêu cầu đã được duyệt.")
            # Refresh the treeview
            tree.delete(*tree.get_children())
            for _, row in df.iterrows():
                tree.insert("", "end", values=(row["ID"], row["Số nhập kho"], row["Mã sách"], row["Người mượn"], row["Thời gian"], row["Trạng thái"]))
        except Exception as e:
            logging.error(f"Lỗi khi duyệt yêu cầu: {str(e)}")
            messagebox.showerror("Lỗi", f"Lỗi khi duyệt yêu cầu: {str(e)}")

    def deny():
        nonlocal df  # Use nonlocal to modify the outer df
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn một yêu cầu để từ chối.")
            return
        values = tree.item(selected[0])['values']
        request_id = values[0]
        status = values[5]
        if status != "Chờ duyệt":
            messagebox.showwarning("Lỗi", f"Yêu cầu này đã được xử lý (Trạng thái: {status}).")
            return
        try:
            # Update the DataFrame
            df.loc[df["ID"] == request_id, "Trạng thái"] = "Từ chối"
            # Save the updated DataFrame to Excel
            df.to_excel("borrow_requests.xlsx", index=False)
            messagebox.showinfo("Đã cập nhật", "Yêu cầu đã bị từ chối.")
            # Refresh the treeview
            tree.delete(*tree.get_children())
            for _, row in df.iterrows():
                tree.insert("", "end", values=(row["ID"], row["Số nhập kho"], row["Mã sách"], row["Người mượn"], row["Thời gian"], row["Trạng thái"]))
        except PermissionError:
            logging.error("Không thể lưu file borrow_requests.xlsx: Permission denied")
            messagebox.showerror("Lỗi", "Không thể lưu file! Vui lòng đóng file Excel nếu đang mở hoặc kiểm tra quyền ghi.")
        except Exception as e:
            logging.error(f"Lỗi khi từ chối yêu cầu: {str(e)}")
            messagebox.showerror("Lỗi", f"Lỗi khi từ chối yêu cầu: {str(e)}")

    def return_book():
        nonlocal df
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn một yêu cầu để xác nhận trả sách.")
            return
        values = tree.item(selected[0])['values']
        request_id = values[0]
        snk = values[1]
        status = values[5]
        if status != "Đã duyệt":
            messagebox.showwarning("Lỗi", f"Chỉ có thể xác nhận trả sách cho yêu cầu đã duyệt (Trạng thái: {status}).")
            return
        try:
            # Update the DataFrame
            df.loc[df["ID"] == request_id, "Trạng thái"] = "Đã trả"
            detail_list.update_status(snk, "Có sẵn")
            # Save the updated DataFrame to Excel
            df.to_excel("borrow_requests.xlsx", index=False)
            # Ghi vào lịch sử trả sách
            if not os.path.exists("borrow_history.xlsx"):
                pd.DataFrame(columns=["Số nhập kho", "Mã sách", "Người mượn", "Thời gian mượn", "Thời gian trả"]).to_excel("borrow_history.xlsx", index=False)
            history_df = pd.read_excel("borrow_history.xlsx")
            history_entry = pd.DataFrame([{
                "Số nhập kho": snk,
                "Mã sách": values[2],
                "Người mượn": values[3],
                "Thời gian mượn": values[4],
                "Thời gian trả": datetime.now().strftime("%d/%m/%Y %H:%M")
            }])
            history_df = pd.concat([history_df, history_entry], ignore_index=True)
            history_df.to_excel("borrow_history.xlsx", index=False)
            messagebox.showinfo("Thành công", "Sách đã được xác nhận trả.")
            # Refresh the treeview
            tree.delete(*tree.get_children())
            for _, row in df.iterrows():
                tree.insert("", "end", values=(row["ID"], row["Số nhập kho"], row["Mã sách"], row["Người mượn"], row["Thời gian"], row["Trạng thái"]))
        except PermissionError:
            logging.error("Không thể lưu file: Permission denied")
            messagebox.showerror("Lỗi", "Không thể lưu file! Vui lòng đóng file Excel nếu đang mở hoặc kiểm tra quyền ghi.")
        except Exception as e:
            logging.error(f"Lỗi khi xác nhận trả sách: {str(e)}")
            messagebox.showerror("Lỗi", f"Lỗi khi xác nhận trả sách: {str(e)}")

    tk.Button(window, text="Duyệt", command=approve).pack(side=tk.LEFT, padx=10, pady=10)
    tk.Button(window, text="Từ chối", command=deny).pack(side=tk.LEFT, padx=10, pady=10)
    tk.Button(window, text="Xác nhận trả sách", command=return_book).pack(side=tk.LEFT, padx=10, pady=10)
# ========================= Nhập/Xuất file Excel =========================

    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if not filepath:
        return
    try:
        df = pd.DataFrame([
            (b.Ma_sach, b.ten_sach, b.tac_gia, b.nha_xuat_ban, b.nam_xuat_ban, b.so_ISBN)
            for b in book_list.books.get_list()
        ], columns=["Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Năm xuất bản", "Số ISBN"])
        df.to_excel(filepath, index=False)
        messagebox.showinfo("Thành công", "Đã lưu dữ liệu vào file Excel.")
    except PermissionError:
        logging.error(f"Không thể lưu file {filepath}: Permission denied")
        messagebox.showerror("Lỗi", "Không thể lưu file! Vui lòng đóng file Excel nếu đang mở hoặc kiểm tra quyền ghi.")
    except Exception as e:
        logging.error(f"Lỗi khi lưu file {filepath}: {str(e)}")
        messagebox.showerror("Lỗi", f"Lỗi khi lưu file: {str(e)}")

# ========================= Giao diện đăng nhập =========================
def authenticate(callback):
    login_window = Toplevel(root)
    login_window.title("Đăng nhập")
    login_window.geometry("400x200")
    
    login_attempts = [0]  # Sử dụng list để thay đổi giá trị trong closure
    max_attempts = 3

    tk.Label(login_window, text="Vai trò:", width=15, anchor="e").grid(row=0, column=0, padx=5, pady=10, sticky="e")
    role_var = tk.StringVar(value="Quản thư")  # Mặc định là Quản thư
    role_menu = ttk.Combobox(login_window, textvariable=role_var, values=["Quản thư", "Sinh viên"], state="readonly", width=20)
    role_menu.grid(row=0, column=1, padx=5, pady=10, sticky="w")
    
    tk.Label(login_window, text="Tên người dùng:", width=15, anchor="e").grid(row=1, column=0, padx=5, pady=10, sticky="e")
    username_entry = tk.Entry(login_window, width=23)
    username_entry.grid(row=1, column=1, padx=5, pady=10, sticky="w")
    
    tk.Label(login_window, text="Mật khẩu:", width=15, anchor="e").grid(row=2, column=0, padx=5, pady=10, sticky="e")
    password_entry = tk.Entry(login_window, show="*", width=23)
    password_entry.grid(row=2, column=1, padx=5, pady=10, sticky="w")

    # Cấu hình để căn giữa
    login_window.grid_columnconfigure(0, weight=1)
    login_window.grid_columnconfigure(1, weight=1)
    login_window.grid_rowconfigure(0, weight=1)
    login_window.grid_rowconfigure(1, weight=1)
    login_window.grid_rowconfigure(2, weight=1)
    login_window.grid_rowconfigure(3, weight=1)

    def check_login():
        current_role = role_var.get()
        username = username_entry.get().strip()
        password = password_entry.get().strip()
        
        if not username or not password:
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ tên người dùng và mật khẩu!")
            return

        authenticated = False
        name = None
        if current_role == "Quản thư":
            authenticated = teacher_list.authenticate(username, password)
        elif current_role == "Sinh viên":
            name = student_list.authenticate(username, password)
            authenticated = name is not False

        if authenticated:
            logging.info(f"Đăng nhập thành công - Vai trò: {current_role}, Username: {username} vào {datetime.now()}")
            login_window.destroy()  # Đóng cửa sổ đăng nhập
            if current_role == "Quản thư":
                callback(current_role, username)
            else:  # Student
                callback(current_role, username, name)
        else:
            login_attempts[0] += 1
            remaining = max_attempts - login_attempts[0]
            logging.warning(f"Đăng nhập thất bại - Vai trò: {current_role}, Username: {username} vào {datetime.now()}")
            if remaining > 0:
                messagebox.showerror("Lỗi", f"Tên người dùng hoặc mật khẩu sai! Còn {remaining} lần thử.")
                username_entry.delete(0, tk.END)
                password_entry.delete(0, tk.END)
            else:
                messagebox.showerror("Lỗi", "Đã vượt quá số lần thử. Vui lòng thử lại sau!")
                login_window.destroy()

    def on_return(event):
        check_login()

    username_entry.bind("<Return>", on_return)
    password_entry.bind("<Return>", on_return)
    
    tk.Button(login_window, text="Đăng nhập", command=check_login, width=10).grid(row=3, column=0, pady=10)
    tk.Button(login_window, text="Thoát", command=lambda: [login_window.destroy(), root.quit()], width=10).grid(row=3, column=1, pady=10)

# ========================= Giao diện chi tiết sách (Teacher) =========================
def open_detail_window():
    detail_window = tk.Toplevel(root)
    detail_window.title("BẢNG THÔNG TIN SÁCH")
    detail_window.geometry("600x450")

    tk.Label(detail_window, text="Số nhập kho").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    tk.Label(detail_window, text="Mã sách").grid(row=1, column=0, padx=5, pady=5, sticky="e")


    entry_snk = tk.Entry(detail_window, width=30)
    entry_ms = tk.Entry(detail_window, width=30)

    entry_snk.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    entry_ms.grid(row=1, column=1, padx=5, pady=5, sticky="w")


    tk.Label(detail_window, text="Tìm kiếm (Số nhập kho/Mã sách/Trạng thái):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
    entry_detail_search = tk.Entry(detail_window, width=30)
    entry_detail_search.grid(row=3, column=1, padx=5, pady=5, sticky="w")

    def search_detail():
        term = entry_detail_search.get().strip().lower()
        if not term:
            messagebox.showerror("Lỗi", "Vui lòng nhập từ khóa tìm kiếm!")
            return
        for row in tree_detail.get_children():
            tree_detail.delete(row)
        curr = detail_list.details.head
        while curr:
            d = curr.data
            if (term in str(d.so_nhap_kho).lower() or term in str(d.ma_sach).lower() or term in str(d.tinh_trang).lower()):
                tree_detail.insert("", "end", values=(d.so_nhap_kho, d.ma_sach, d.tinh_trang))
            curr = curr.next

    tk.Button(detail_window, text="Tìm", command=search_detail).grid(row=3, column=2, padx=5, pady=5)
    tk.Button(detail_window, text="\u21BB", command=lambda: detail_list.update_detail_treeview()).grid(row=3, column=3, padx=5, pady=5)

    def add_detail():
        snk = entry_snk.get().strip()
        ms = entry_ms.get().strip()
        if not (snk and ms):
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!")
            return
        detail = BookDetail(snk, ms, "Có sẵn")
        detail_list.add_detail(detail)
        entry_snk.delete(0, tk.END)
        entry_ms.delete(0, tk.END)

    def delete_detail():
        sel = tree_detail.selection()
        if not sel:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn một sách để xóa.")
            return
        snk = tree_detail.item(sel[0])['values'][0]
        detail_list.delete_detail(snk)



    tk.Button(detail_window, text="Thêm sách", command=add_detail).grid(row=4, column=0, padx=5, pady=10)
    tk.Button(detail_window, text="Xóa sách", command=delete_detail).grid(row=4, column=1, padx=5, pady=10)

    cols = ["Số nhập kho", "Mã sách", "Trạng thái"]
    global tree_detail
    tree_detail = ttk.Treeview(detail_window, columns=cols, show="headings")
    for col in cols:
        tree_detail.heading(col, text=col)
        tree_detail.column(col, width=150)
    detail_window.grid_rowconfigure(5, weight=1)
    for c in range(4):
        detail_window.grid_columnconfigure(c, weight=1)
    tree_detail.grid(row=5, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")
    vsb = ttk.Scrollbar(detail_window, orient="vertical", command=tree_detail.yview)
    vsb.grid(row=5, column=4, sticky="ns")
    tree_detail.configure(yscrollcommand=vsb.set)
    detail_list.update_detail_treeview()

# ========================= Giao diện sinh viên =========================
def student_interface(role, username, name):
    # Ẩn cửa sổ chính trước khi mở giao diện sinh viên
    root.withdraw()

    student_window = tk.Toplevel(root)
    student_window.title(f"TRA CỨU & MƯỢN SÁCH - {name}")
    student_window.geometry("600x400")

    tk.Label(student_window, text="Danh sách sách có thể mượn").pack(pady=5)
    student_tree = ttk.Treeview(student_window, columns=["Mã sách", "Tên sách"], show="headings")
    student_tree.heading("Mã sách", text="Mã sách")
    student_tree.heading("Tên sách", text="Tên sách")
    student_tree.column("Mã sách", width=150)
    student_tree.column("Tên sách", width=300)
    student_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    vsb = ttk.Scrollbar(student_window, orient="vertical", command=student_tree.yview)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    student_tree.configure(yscrollcommand=vsb.set)

    def load_books():
        for row in student_tree.get_children():
            student_tree.delete(row)
        current = book_list.books.head
        while current:
            b = current.data
            student_tree.insert("", "end", values=(b.Ma_sach, b.ten_sach))
            current = current.next

    def on_select(event):
        selected = student_tree.selection()
        if not selected:
            return
        ma_sach = student_tree.item(selected[0])['values'][0]
        detail_window = tk.Toplevel(student_window)
        detail_window.title(f"Chi tiết trạng thái - {ma_sach}")
        detail_window.geometry("400x300")

        tree = ttk.Treeview(detail_window, columns=["Số nhập kho", "Trạng thái"], show="headings")
        tree.heading("Số nhập kho", text="Số nhập kho")
        tree.heading("Trạng thái", text="Trạng thái")
        tree.column("Số nhập kho", width=150)
        tree.column("Trạng thái", width=150)
        tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        vsb = ttk.Scrollbar(detail_window, orient="vertical", command=tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=vsb.set)

        def load_detail():
            for row in tree.get_children():
                tree.delete(row)
            curr = detail_list.details.head
            while curr:
                d = curr.data
                if d.ma_sach == ma_sach:
                    tree.insert("", "end", values=(d.so_nhap_kho, d.tinh_trang))
                curr = curr.next

        def request_borrow():
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("Chưa chọn", "Vui lòng chọn sách còn Có sẵn để mượn")
                return
            snk = tree.item(selected_item[0])['values'][0]
            tt = tree.item(selected_item[0])['values'][1]
            if tt != "Có sẵn":
                messagebox.showwarning("Không hợp lệ", "Sách này đang được mượn")
                return
            request_id = str(uuid.uuid4())
            borrower_name = student_list.get_student_name(username)
            if not os.path.exists("borrow_requests.xlsx"):
                pd.DataFrame(columns=["ID", "Số nhập kho", "Mã sách", "Người mượn", "Thời gian", "Trạng thái"]).to_excel("borrow_requests.xlsx", index=False)
            try:
                df = pd.read_excel("borrow_requests.xlsx")
                new_request = pd.DataFrame([{
                    "ID": request_id,
                    "Số nhập kho": snk,
                    "Mã sách": ma_sach,
                    "Người mượn": borrower_name,
                    "Thời gian": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "Trạng thái": "Chờ duyệt"
                }])
                df = pd.concat([df, new_request], ignore_index=True)
                df.to_excel("borrow_requests.xlsx", index=False)
                messagebox.showinfo("Thành công", "Yêu cầu đã gửi. Chờ quản thư duyệt.")
            except Exception as e:
                logging.error(f"Lỗi khi gửi yêu cầu mượn: {str(e)}")
                messagebox.showerror("Lỗi", f"Lỗi khi gửi yêu cầu: {str(e)}")

        def view_requests():
            request_window = Toplevel(detail_window)
            request_window.title("Trạng thái yêu cầu mượn sách")
            request_window.geometry("600x300")
            tree_requests = ttk.Treeview(request_window, columns=["ID", "Số nhập kho", "Mã sách", "Người mượn", "Thời gian", "Trạng thái"], show="headings")
            for col in ["ID", "Số nhập kho", "Mã sách", "Người mượn", "Thời gian", "Trạng thái"]:
                tree_requests.heading(col, text=col)
                tree_requests.column(col, width=100)
            tree_requests.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            vsb = ttk.Scrollbar(request_window, orient="vertical", command=tree_requests.yview)
            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            tree_requests.configure(yscrollcommand=vsb.set)
            if os.path.exists("borrow_requests.xlsx"):
                try:
                    df = pd.read_excel("borrow_requests.xlsx")
                    for _, row in df.iterrows():
                        if row["Mã sách"] == ma_sach and row["Người mượn"] == name:
                            tree_requests.insert("", "end", values=(row["ID"], row["Số nhập kho"], row["Mã sách"], row["Người mượn"], row["Thời gian"], row["Trạng thái"]))
                except Exception as e:
                    logging.error(f"Lỗi khi xem yêu cầu: {str(e)}")
                    messagebox.showerror("Lỗi", f"Lỗi khi xem yêu cầu: {str(e)}")

        tk.Button(detail_window, text="Yêu cầu mượn sách", command=request_borrow).pack(pady=5)
        tk.Button(detail_window, text="Xem trạng thái yêu cầu", command=view_requests).pack(pady=5)
        load_detail()

    # Hàm tra cứu danh sách yêu cầu mượn sách của sinh viên
    def view_all_requests():
        request_window = Toplevel(student_window)
        request_window.title("Danh sách yêu cầu mượn sách")
        request_window.geometry("600x400")

        tree_requests = ttk.Treeview(request_window, columns=["ID", "Số nhập kho", "Mã sách", "Thời gian", "Trạng thái"], show="headings")
        for col in ["ID", "Số nhập kho", "Mã sách", "Thời gian", "Trạng thái"]:
            tree_requests.heading(col, text=col)
            tree_requests.column(col, width=120)
        tree_requests.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        vsb = ttk.Scrollbar(request_window, orient="vertical", command=tree_requests.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        tree_requests.configure(yscrollcommand=vsb.set)

        def load_requests():
            for row in tree_requests.get_children():
                tree_requests.delete(row)
            if not os.path.exists("borrow_requests.xlsx"):
                messagebox.showinfo("Thông báo", "Bạn chưa có yêu cầu mượn sách nào.")
                return
            try:
                df = pd.read_excel("borrow_requests.xlsx")
                for _, row in df.iterrows():
                    if row["Người mượn"] == name:
                        tree_requests.insert("", "end", values=(row["ID"], row["Số nhập kho"], row["Mã sách"], row["Thời gian"], row["Trạng thái"]))
            except Exception as e:
                logging.error(f"Lỗi khi xem danh sách yêu cầu: {str(e)}")
                messagebox.showerror("Lỗi", f"Lỗi khi xem danh sách yêu cầu: {str(e)}")


        load_requests()

    student_tree.bind("<Double-1>", on_select)
    tk.Button(student_window, text="Tra cứu yêu cầu mượn sách", command=view_all_requests).pack(pady=5)
    tk.Button(student_window, text="Đăng xuất", command=lambda: [student_window.destroy(), authenticate(lambda r, u, n=None: teacher_interface(r, u) if r == "Quản thư" else student_interface(r, u, n))]).pack(pady=5)
    load_books()

# ========================= Giao diện giáo viên =========================
def teacher_interface(role, username):
    root.deiconify()  # Hiển thị lại cửa sổ chính
    # Xóa nội dung cũ trên root nếu có
    for widget in root.winfo_children():
        widget.destroy()

    main_frame = tk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    main_frame.grid_rowconfigure(0, weight=0)
    main_frame.grid_rowconfigure(1, weight=0)
    main_frame.grid_rowconfigure(2, weight=0)
    main_frame.grid_rowconfigure(3, weight=1)
    main_frame.grid_rowconfigure(4, weight=0)
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    main_frame.grid_columnconfigure(2, weight=1)

    labels = ["Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Năm xuất bản", "Số ISBN"]
    entries = []
    for i, text in enumerate(labels):
        tk.Label(main_frame, text=text + ":", width=20, anchor='w').grid(row=i, column=0, padx=10, pady=5)
        entry = tk.Entry(main_frame, width=30)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries.append(entry)
    def sort_data_by_column(column_index):
     # Lấy dữ liệu từ Treeview
        data = [(tree.set(child, "#1"), tree.set(child, "#2"), tree.set(child, "#3"), 
                tree.set(child, "#4"), tree.set(child, "#5"), tree.set(child, "#6")) 
                for child in tree.get_children()]
    # Hàm hỗ trợ để so sánh giá trị, đặc biệt cho cột "Mã sách"
        def extract_number(ma):
            so = ''.join(ch for ch in str(ma) if ch.isdigit())
            return int(so) if so else 0
        # Bubble sort
        n = len(data)
        for i in range(n):
            swapped = False
            for j in range(0, n - i - 1):
                # So sánh dựa trên column_index
                value1 = data[j][column_index]
                value2 = data[j + 1][column_index]
                # Đặc biệt: nếu sắp xếp theo cột "Mã sách" (column_index=0), dùng extract_number
                if column_index == 0:
                    value1 = extract_number(value1)
                    value2 = extract_number(value2)
                # Đặc biệt: nếu sắp xếp theo cột "Năm xuất bản" (column_index=4), chuyển thành số
                elif column_index == 4:
                    value1 = int(value1) if value1.isdigit() else 0
                    value2 = int(value2) if value2.isdigit() else 0 
                # So sánh và hoán đổi
                if value1 > value2:
                    data[j], data[j + 1] = data[j + 1], data[j]
                    swapped = True
            if not swapped:
                break
        # Xóa các dòng cũ trong Treeview
        tree.delete(*tree.get_children())
        # Thêm các dòng đã sắp xếp
        for item in data:
            tree.insert("", tk.END, values=item)
    def sort_data_window():
        sort_window = tk.Toplevel(root)
        sort_window.title("Chọn thuộc tính sắp xếp")
        sort_window.geometry("300x300")
        tk.Label(sort_window, text="Chọn thuộc tính để sắp xếp:").pack(pady=10)
        def sort_by(column):
            try:
                if column == "Mã sách":
                    sort_data_by_column(0)
                elif column == "Tên sách":
                    sort_data_by_column(1)
                elif column == "Tác giả":
                    sort_data_by_column(2)
                elif column == "Nhà xuất bản":
                    sort_data_by_column(3)
                elif column == "Năm xuất bản":
                    sort_data_by_column(4)
                elif column == "Số ISBN":
                    sort_data_by_column(5)
                sort_window.destroy()
            except Exception as e:
                logging.error(f"Lỗi khi sắp xếp theo {column}: {str(e)}")
                messagebox.showerror("Lỗi", f"Lỗi khi sắp xếp: {str(e)}")
        options = ["Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Năm xuất bản", "Số ISBN"]
        for option in options:
            tk.Button(sort_window, text=option, command=lambda opt=option: sort_by(opt)).pack(fill="both", padx=10, pady=5)
    def import_excel():
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not filepath:
            return
        try:
            df = pd.read_excel(filepath)
            expected_columns = ["Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Năm xuất bản", "Số ISBN"]
            if not all(col in df.columns for col in expected_columns):
                messagebox.showerror("Lỗi", "File Excel không đúng định dạng!")
                return
            duplicates = []
            current = book_list.books.head
            existing_ma_sach = set()
            while current:
                existing_ma_sach.add(current.data.Ma_sach)
                current = current.next
            for _, row in df.iterrows():
                if row["Mã sách"] in existing_ma_sach:
                    duplicates.append(row["Mã sách"])
                else:
                    book = Book(
                        row["Mã sách"], row["Tên sách"], row["Tác giả"],
                        row["Nhà xuất bản"], row["Năm xuất bản"], row["Số ISBN"]
                    )
                    if not row["Năm xuất bản"] or not str(row["Năm xuất bản"]).isdigit() or int(row["Năm xuất bản"]) < 0:
                        messagebox.showerror("Lỗi", f"Năm xuất bản không hợp lệ tại Mã sách: {row['Mã sách']}")
                        return
                    if not re.match(r"^[a-zA-Z0-9\s]+$", str(row["Mã sách"])):
                        messagebox.showerror("Lỗi", f"Mã sách không hợp lệ tại: {row['Mã sách']}")
                        return
                    if not re.match(r"^\d{10}|\d{13}$", str(row["Số ISBN"]).replace("-", "")):
                        messagebox.showerror("Lỗi", f"Số ISBN không hợp lệ tại Mã sách: {row['Mã sách']}")
                        return
                    book_list.books.appendLast(book)
            if duplicates:
                messagebox.showerror("Lỗi", f"Các Mã sách đã tồn tại: {', '.join(duplicates)}")
                return
            book_list.sort_by_ma_sach()
            book_list.save_to_excel()
            update_treeview(book_list.books)
            messagebox.showinfo("Thành công", "Đã nhập dữ liệu từ file Excel.")
        except Exception as e:
            logging.error(f"Lỗi khi nhập file: {str(e)}")
            messagebox.showerror("Lỗi", f"Lỗi khi nhập file: {str(e)}")

    def save_excel():
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not filepath:
            return
        try:
            df = pd.DataFrame([
                (b.Ma_sach, b.ten_sach, b.tac_gia, b.nha_xuat_ban, b.nam_xuat_ban, b.so_ISBN)
                for b in book_list.books.get_list()
            ], columns=["Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Năm xuất bản", "Số ISBN"])
            df.to_excel(filepath, index=False)
            messagebox.showinfo("Thành công", "Đã lưu dữ liệu vào file Excel.")
        except PermissionError:
            logging.error(f"Không thể lưu file {filepath}: Permission denied")
            messagebox.showerror("Lỗi", "Không thể lưu file! Vui lòng đóng file Excel nếu đang mở hoặc kiểm tra quyền ghi.")
        except Exception as e:
            logging.error(f"Lỗi khi lưu file {filepath}: {str(e)}")
            messagebox.showerror("Lỗi", f"Lỗi khi lưu file: {str(e)}")
    global entry_code, entry_title, entry_author, entry_publisher, entry_year, entry_isbn, entry_search
    entry_code, entry_title, entry_author, entry_publisher, entry_year, entry_isbn = entries

    tk.Button(main_frame, text="Thêm mới biên mục sách", command=book_list.add_book).grid(row=len(labels), column=0, pady=10)
    tk.Button(main_frame, text="Sửa biên mục sách", command=book_list.edit_book).grid(row=len(labels), column=1, pady=10)
    tk.Button(main_frame, text="Xóa biên mục sách", command=book_list.delete_book).grid(row=len(labels), column=2, pady=10)

    tk.Label(main_frame, text="Tìm kiếm (Mã sách/Tên sách/Số ISBN):", anchor="w").grid(row=len(labels)+1, column=0, padx=5, pady=5, sticky="w")
    entry_search = tk.Entry(main_frame, width=30)
    entry_search.grid(row=len(labels)+1, column=1, padx=5, pady=5, sticky="w")
    tk.Button(main_frame, text="Tìm kiếm", command=book_list.search_book_multi).grid(row=len(labels)+1, column=2, padx=5, pady=5, sticky="w")
    tk.Button(main_frame, text="\u21BB", command=lambda: [entry_search.delete(0, tk.END), update_treeview(book_list.books)], width=2).grid(row=len(labels)+1, column=3, padx=5, pady=5, sticky="w")

    cols = ["Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Năm xuất bản", "Số ISBN"]
    global tree
    tree = ttk.Treeview(main_frame, columns=cols, show="headings")
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=120)
    tree.grid(row=len(labels)+2, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")
    vsb = ttk.Scrollbar(main_frame, orient="vertical", command=tree.yview)
    vsb.grid(row=len(labels)+2, column=4, sticky="ns")
    tree.configure(yscrollcommand=vsb.set)

    tk.Button(main_frame, text="Duyệt yêu cầu mượn sách", command=approve_requests, width=10).grid(row=len(labels)+3, column=0, padx=10, pady=5, sticky="ew")
    tk.Button(main_frame, text="Sắp xếp", command=sort_data_window, width=10).grid(row=len(labels)+4, column=0, padx=10, pady=5, sticky="ew")
    tk.Button(main_frame, text="Thông tin chi tiết sách", command=open_detail_window, width=10).grid(row=len(labels)+5, column=0, padx=10, pady=5, sticky="ew")
    tk.Button(main_frame, text="Nhập file", command=import_excel, width=20).grid(row=len(labels)+3, column=2, padx=10, pady=5, sticky="ew")
    tk.Button(main_frame, text="Xuất file", command=save_excel, width=20).grid(row=len(labels)+4, column=2, padx=10, pady=5, sticky="ew")
    tk.Button(main_frame, text="Thoát khỏi hệ thống", command=lambda: [root.withdraw(), authenticate(lambda r, u: teacher_interface(r, u) if r == "Quản thư" else student_interface(r, u))], width=20).grid(row=len(labels)+5, column=2, padx=10, pady=5, sticky="ew")

    update_treeview(book_list.books)

# ========================= Chạy chương trình =========================
root = tk.Tk()
root.title("CHƯƠNG TRÌNH QUẢN LÝ THƯ VIỆN")
root.withdraw()

# Khởi tạo tree_detail mặc định
tree_detail = None

base_dir = os.path.dirname(os.path.abspath(__file__))
book_list = BookList(os.path.join(base_dir, "book_data.xlsx"))
detail_list = BookDetailList(os.path.join(base_dir, "book_inf.xlsx"))
teacher_list = TeacherList(os.path.join(base_dir, "teachers.xlsx"))
student_list = StudentList(os.path.join(base_dir, "students.xlsx"))

authenticate(lambda role, username, name=None: teacher_interface(role, username) if role == "Quản thư" else student_interface(role, username, name))
root.mainloop()