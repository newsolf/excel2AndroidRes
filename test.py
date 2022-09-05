if __name__ == '__main__':
    file_strings = open("outputs/res/values-zh_CN/strings.xml", "w+")
    file_strings.write("new w")
    file_strings.flush()
    file_strings.close()
