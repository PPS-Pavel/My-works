import win32com.client  
import re  

def main():
   
    application = win32com.client.Dispatch("Renga.Application.1")
    application.Visible = True
    application.Enabled = True
    project = application.Project  
    model = project.Model  

    
    roomType = '{f1a805ff-573d-f46b-ffba-57f4bccaa6ed}'
    room_number_param = '{b44d9acd-bd51-4ca4-96a8-2a3f92a685b5}'
    room_name_param = '{aa18be09-18ef-43ec-b89d-3b245874dc39}'
    levelName = '{1bb1addf-a3c0-4356-9525-107ea7df1513}'
    levelType = '{C3CE17FF-6F28-411F-B18D-74FE957B2BA8}'
    levelId = '{8cdf2e5b-03f7-4101-9b43-93b9da18f411}'
    
    unique_room_numbers = {}  
    level_numbers = {}  

    objectCollection = model.GetObjects()  
    
    
    for i in range(objectCollection.Count):
        obj = objectCollection.GetByIndex(i)
        if obj.ObjectTypeS == levelType:  
            paramcontainer = obj.GetParameters()  
            LevelNameValue = paramcontainer.GetS(levelName)  
            namelevel = LevelNameValue.GetStringValue()  

            
            match = re.search(r'Э(\d+)', namelevel)
            if match:
                level_number = match.group(1)  
                level_numbers[obj.Id] = level_number  

    
    for i in range(objectCollection.Count):
        obj = objectCollection.GetByIndex(i)
        if obj.ObjectTypeS.lower() == roomType:  
            param_container = obj.GetParameters()  
            level = obj.GetParameters().GetS(levelId)  
            room_level_id = level.GetIntValue()  
            room_name = param_container.GetS(room_name_param)  
            name_room = room_name.GetStringValue()  
            room_number = param_container.GetS(room_number_param)  
            number_room = room_number.GetStringValue()  
            
            if not number_room:  
                print(f"Комната без номера: {name_room}")

            
            if number_room:
                if number_room in unique_room_numbers:
                    unique_room_numbers[number_room].append(name_room)  
                else:
                    unique_room_numbers[number_room] = [name_room]  

            
            level_number = level_numbers.get(room_level_id)
            
            if number_room and level_number:
                first_number_room = number_room[0]  
                if first_number_room != level_number:  
                    print(f"Номер комнаты '{number_room}' не соответствует этажу '{level_number}' для комнаты '{name_room}'.")

    
    for number in unique_room_numbers:
        if len(unique_room_numbers[number]) > 1:  
            print(f"Комнаты с одинаковым номером '{number}': {', '.join(unique_room_numbers[number])}")

if __name__ == '__main__':
    main()  