# encoding: utf-8
import ComputeExcle as ce
import random

def ChooseData(money,data):
    '''
    每张发票可以1~3个商品
    '''
    good = random.randint(1,3)
    ticket = []
    buildticket(int(money),good,data,ticket)
    return ticket

def buildticket(money,i,data,ticket):
    money = int(money)
    if i == 0:
        return 
    while True:
        b1 = data.popitem()
        # print(b1[1][1],b1[1][0])
        if b1[1][1] > money and b1[1][0] < money:#添加一件
            ticket.append((b1[0],money,1))
            return
        if b1[1][1] < money:
            price = random.randint(b1[1][0],b1[1][0])
            num = random.randint(1,money//price)
            ticket.append((b1[0],price,num))
            buildticket(money-price*num,i-1,data,ticket)
            break
        data[b1[0]] = b1[1]

def face():
    work = []
    data = ce.read_excle("database.xlsx")
    while True:
        tick = input("请输入订单号：")
        money = int(input("请输入金额："))
        detail = ChooseData(money,data)
        work.append((tick,money,detail))
        way = input("是否继续添加订单 y/n\n")
        if way != "y" and way != "Y":
            break
    name = input("请输入保存文件名，默认为data\n")
    if name == "":
        name = "data"
    ce.wirte_excle(name+".xls",work)



if __name__ == "__main__":
    face()
    # print (ChooseData(money,data))

        
