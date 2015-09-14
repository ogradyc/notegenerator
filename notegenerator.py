"""Generate a template for notes for easy note taking in classes."""
#notegenerator/0.1 by mobyte
import datetime
import os
import json

def main():
    """Main function"""
    today = datetime.date.today()
    year = today.year
    month = today.month
    day = today.day
    print "Today's Date: {}-{}-{}".format(month, day, year)
    if os.path.isfile("config.json") == True:
        with open('config.json') as config:
            data = json.load(config)
        config.close()
        print "Current text editor: {}".format(data['editor'])
        menu()
    else:
        setup()
        main()

def setup():
    """Manages config file for the user"""
    data = {'editor': '', 'classes': [], 'classdir': []}
    print "What is your text editor?"
    texteditor = raw_input("> ")
    print "How many classes do you have?"
    numberclasses = raw_input("> ")
    classlist = []
    noteslist = []
    for index in range(0, int(numberclasses)):
        print "What is class {}?".format(index+1)
        userclass = raw_input("> ")
        classlist.append(userclass)
        while True:
            print "Where should the notes for {} be saved?".format(classlist[index])
            noteslocation = raw_input("> ")
            if os.path.exists(noteslocation):
                noteslist.append(noteslocation)
                break
            else:
                print "Invalid location."
    data['editor'] = texteditor
    data['classes'] = classlist
    data['classdir'] = noteslist
    with open('config.json', 'w') as config:
        json.dump(data, config)
    config.close()

def menu():
    """Main menu for program"""
    print "Options:"
    print "1. Take notes"
    print "2. Reset config file"
    print "0. Quit program"
    userinput = raw_input("> ")
    if userinput == "1":
        notes()
    elif userinput == "2":
        os.remove("config.json")
        main()
    elif userinput == "0":
        quit()

def notes():
    """Generates the file for taking notes"""
    print "Which class are you taking notes for?"
    with open('config.json') as config:
        data = json.load(config)
    config.close()
    numberofclasses = len(data["classes"])
    for index in range(0, int(numberofclasses)):
        print "{}. {}".format(index+1, data['classes'][index])
    print "0. Cancel"
    userinput = raw_input("> ")
    if userinput == "0":
        menu()
    else:
        currentclass = data['classes'][int(userinput)-1]
        print "Generating notes file for {}...".format(currentclass)
        today = datetime.date.today()
        date = "{}-{}-{}".format(today.month, today.day, today.year)
        texteditor = data['editor']
        path = data['classdir'][int(userinput)-1]
        filename = "{} [{}]".format(date, currentclass)
        path = "{}/{}".format(path, filename)
        if os.path.isfile(path) == True:
            print "Notes file already exists.\nOverwrite, open, or cancel?"
            print "1. Overwrite"
            print "2. Open"
            print "0. Cancel"
            userinput = raw_input("> ")
            if userinput == "1":
                print "Are you sure you want to overwrite?"
                print "1. Overwrite"
                print "0. Cancel"
                userinput = raw_input("> ")
                if userinput == "1":
                    print "Overwriting..."
                    path = path.replace(" ", "\\ ")
                    template = "{}\n\n{}\n\nAssignments:\n\nNotes:\n".format(currentclass, date)
                    path = path.replace("\\ ", " ")
                    with open(path, 'w+') as note:
                        note.write(template)
                    note.close()
                    print "Successful. Opening {}...".format(texteditor)
                    path = path.replace(" ", "\\ ")
                    command = "{} {}".format(texteditor, path)
                    os.system(command)
                elif userinput == "0":
                    menu()
            elif userinput == "2":
                print "Opening with {}...".format(texteditor)
                path = path.replace(" ", "\\ ")
                command = "{} {}".format(texteditor, path)
                os.system(command)
            elif userinput == "0":
                menu()
        else:
            path = path.replace(" ", "\\ ")
            os.system('touch {}'.format(path))
            path = path.replace("\\ ", " ")
            template = "{}\n\n{}\n\nAssignments:\n\nNotes:\n".format(currentclass, date)
            path = path.replace("\\ ", " ")
            with open(path, 'w+') as note:
                note.write(template)
            note.close()
            print "Successful. Opening {}...".format(texteditor)
            path = path.replace(" ", "\\ ")
            command = "{} {}".format(texteditor, path)
            os.system(command)

if __name__ == "__main__":
    main()
