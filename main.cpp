#include <fstream>
#include <string>
#include <stdio.h>
#include <io.h>
#include <algorithm>
#include <vector> 
#include <iostream>
using namespace std;


void getFiles(string path, vector<string>& files)
{
    //�ļ����
    long   hFile = 0;
    //�ļ���Ϣ
    struct _finddata_t fileinfo;
    string p;
    if ((hFile = _findfirst(p.assign(path).append("\\*").c_str(), &fileinfo)) != -1)
    {
        do
        {
            //�����Ŀ¼,����֮
            //�������,�����б�
            if ((fileinfo.attrib &  _A_SUBDIR))
            {
                if (strcmp(fileinfo.name, ".") != 0 && strcmp(fileinfo.name, "..") != 0)
                    getFiles(p.assign(path).append("\\").append(fileinfo.name), files);
            }
            else
            {
                string str = fileinfo.name;
                str = str.substr(str.rfind("."), str.length());
                //transform(str.begin(), str.end(), str.begin(), tolower);
                if (str == ".ppt" || str == ".pptx")
                {
                    files.push_back(p.assign(path).append("\\").append(fileinfo.name));
                }
            }
        } while (_findnext(hFile, &fileinfo) == 0);
        _findclose(hFile);
    }
}

string&   replace_all(string&   str, const   string&   old_value, const   string&   new_value)
{
    while (true) {
        string::size_type   pos(0);
        if ((pos = str.find(old_value)) != string::npos)
            str.replace(pos, old_value.length(), new_value);
        else   break;
    }
    return   str;
}

int main(int argc, char** argv)
{

    cout << "������PPT����·��:" << endl;
    char input[1000];
    cin.getline(input, sizeof(input));

    string file = input;

    vector<string> vec;
    vec.reserve(100);
    getFiles(file, vec);

    cout << "��Ŀ¼���ܹ���" << vec.size()<<"��PPT/PPTX�ļ�"<<endl;

    for (size_t i = 0; i < vec.size(); ++i)
    {
        char sz[10] = { 0 };

        _itoa_s(i, sz, 10);

        string genPath = string("D:\\") + sz;

        cout << "��" + i << "���ļ�,�ļ�·��:" + vec[i] << endl;

        string cmd = string("7z x ") + "\"" + vec[i] + "\" -o" + genPath;

        cout << "��ѹ��" << endl;

        system(cmd.c_str());

        cout << "�������" << endl;
        string filename = vec[i].substr(vec[i].rfind("\\") + 1, vec[i].rfind(".") + 1);

        filename = filename.substr(0, filename.rfind("."));

        string vvv = genPath + string("\\ppt\\presentation.xml");
        ifstream in(vvv.c_str());

        istreambuf_iterator<char> beg(in), end;

        string str(beg, end);

        size_t left = str.find("<p:modifyVerifier");
        if (left == size_t(-1))
        {
            cout << filename<<"PPT�������!!!" << endl;
            cmd = "copy  \"" + vec[i] + "\"" + " D:\\Temp";
            system(cmd.c_str());
            continue;
        }

        string strLeft = str.substr(0, left);

        string temp = str.substr(left, str.length());

        string strRight = temp.substr(temp.find("/>") + 2, temp.length());

        string total = strLeft + strRight;


        vvv = genPath + string("\\ppt\\presentation.xml");
        ofstream out(vvv.c_str());

        out.write(total.c_str(), total.size());

        out.flush();

        out.close();

        in.close();

        cout << "�������!!!" << endl;

        replace_all(filename, "?", " ");

        cmd = "7z a \"d:\\temp\\" + filename + ".zip\" " + genPath + "\\*";

        system(cmd.c_str());

        cout << "ѹ�����!!!" << endl;

        string old = string("D:\\temp\\") + filename + string(".zip");

        string newfile = string("D:\\temp\\") + filename + string(".pptx");

        rename(old.c_str(), newfile.c_str());

        remove(old.c_str());

        string tmp = string("rd /s/q ") + genPath;

        system(tmp.c_str());

        cout << "ɾ������" << endl;

        cout << "�ļ���:" << newfile << endl;
    }

    cout << "ȫ��ת������� ^-^ " << endl;

    system("pause");

    return 0;
}

