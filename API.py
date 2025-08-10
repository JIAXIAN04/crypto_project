# print("hello")
# print("ggg")
# user_weight = float(input("請輸入你的體重: "))
# user_height = float(input("請輸入你的身高: "))
# user_BMI = user_weight / (user_height)**2
# print("你的BMI為: " + str(user_BMI))
#
# if user_BMI<= 18.5:
#     print("此BMI屬於偏瘦")
# elif 18.5 < user_BMI <=25:
#     print("此BMI屬於正常")
# elif 25 < user_BMI <=30:
#     print("此BMI屬於篇胖")
# else:
#     print("你好胖")

#list and append
shopping_list =["apple", "banana"]
shopping_list.append("milk")
print(len(shopping_list))
shopping_list[1]="TV"
print(shopping_list)

num_list = [1, 13, -5, 8, 95]
print(max(num_list))
print(min(num_list))
print(sorted(num_list))