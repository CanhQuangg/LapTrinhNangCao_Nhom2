class Person():
    def __init__(self,id, name, gender, adress, date, phoneNumber):
        self.id = id
        self.name = name
        self.gender = gender
        self.date = date
        self.address = adress
        self.phoneNumber = phoneNumber

    def getInfo(self):
        self.info = [self.id, self.name, self.gender, self.date, self.address, self.phoneNumber]
        return self.info

class Staff(Person):
    def __init__(self, id, name, gender, date, address, phoneNumber, role):
        super().__init__(id, name, gender, address, date, phoneNumber)
        self.role = role

    def getInfo(self):
        self.info = [self.id,
                     self.name,
                     self.gender,
                     self.date,
                     self.address,
                     self.phoneNumber,
                     self.role]
        return self.info

class Customer(Person):
    def __init__(self, id, name, gender, date, address, phoneNumber):
        super().__init__(id, name, gender, address, date, phoneNumber)

    def getInfo(self):
        self.info = [self.id,
                     self.name,
                     self.gender,
                     self.date,
                     self.address,
                     self.phoneNumber]
        return self.info

class Product():
    def __init__(self, id, name, price):
        self.id = id
        self.name = name
        self.price = price

    def getInfo(self):
        self.info = [self.id, self.name, self.price]
        return self.info

class Material():
    def __init__(self,id, name, quantity, unit):
        self.id = id
        self.name = name
        self.quantity = quantity
        self.unit = unit

    def getInfo(self):
        self.info = [self.id, self.name, self.quantity, self.unit]
        return self.info

class Bill():
    def __init__(self,billID, cusID, sellerID, total, date):
        self.billID = billID
        self.cusID = cusID
        self.sellerID = sellerID
        self.total = total
        self.date = date

    def getInfo(self):
        self.info = [self.billID, self.cusID, self.sellerID, self.total, self.date]
        return self.info