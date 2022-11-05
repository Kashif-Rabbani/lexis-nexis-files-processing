# This file is an alternative of python notebook.
import docx2python
from os import listdir
import csv
from os.path import isfile, join, isdir

folder_path = "Air Pollution Data"
file_paths = []


def read_word_file(address, writer, index_to_start_from):
    raw_document = docx2python.docx2python(address)
    content = raw_document.body_runs
    parse_content(content[index_to_start_from][0][0], writer)


def parse_content(content, writer):
    row = []
    for counter in range(0, len(content)):
        if len(content[counter]) > 0:
            if "<a href=\"http" in content[counter][0]:
                title = content[counter][0]
                title = (title.split(">")[1].split("<"))[0]
                row.append(title)

            if content[counter][0] == 'Body':
                # get content now
                local_counter = counter + 1
                body_content_indices = []

                while len(content[local_counter]) == 0:
                    local_counter += 1

                while "\nLoad-Date" not in content[local_counter][0]:
                    if len(content[local_counter]) > 0:
                        body_content_indices.append(local_counter)
                    local_counter += 1
                    while len(content[local_counter]) == 0:
                        local_counter += 1
                text = ""
                for para in body_content_indices:
                    for element in content[para]:
                        text = join(text, element)
                row.append(text)
                writer.writerow(row)
                row = []


def init_reader_writer(file_path, index_to_start_from):
    # open the file in the write mode
    f = open(file_path + ".csv", 'w')
    # create the csv writer
    writer = csv.writer(f)
    header = ['headline', 'body']
    # write a row to the csv file
    writer.writerow(header)
    read_word_file(file_path, writer, index_to_start_from)
    # close the file
    f.close()


def get_files(path):
    for f in listdir(path):
        if isfile(join(path, f)):
            # print("file " + join(path, f))
            if (f.endswith(".DOCX")):
                file_paths.append(join(path, f))

        if isdir(join(path, f)):
            # print("dir " +  join(directory, f))
            get_files(join(path, f))
    return file_paths


def read_files_from_dir(path):
    files = get_files(path)
    print(files)
    for file_path in files:
        file_name = file_path.split("/")[-1]
        limits = file_name[file_name.find("(") + 1:file_name.find(")")]
        index_to_start_from = limits.split("-")[1]
        print("Reading " + file_path)
        remainder = int(index_to_start_from) % 100
        if remainder == 0:
            # print(index_to_start_from)
            init_reader_writer(file_path, 200)
        else:
            index_to_start_from = remainder
            # print(index_to_start_from)
            init_reader_writer(file_path, int(index_to_start_from) * 2)


if __name__ == '__main__':
    read_files_from_dir(folder_path)
    # init_reader_writer("Air Pollution Data/Bangladesh/The Financial Express (106 RESULTS)/Files (1-100).DOCX", 200)
