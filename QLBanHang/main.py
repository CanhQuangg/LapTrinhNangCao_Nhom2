from GUI import *

"""
    danh sách tài khoản đăng nhập:
        manager:
            MN01	    123456
            MN02	    123456
            CanhQuang	15112001
            q           q    
        staff:
            ST01	123456
            ST02	123456
            ST03	123456
            ST04	123456
    
    Lưu ý: chỉ có thể đăng nhập với vai trò quản lý "manager"
"""

# tạo đối tượng cửa sổ hiển thị
window = Tk()

# tạo đối tượng giao diện đăng nhập
GUI_Login(window)
