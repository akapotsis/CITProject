#include <iostream>
#include "include/dirent.h"
#include <errno.h>
#include <string>
#include <tchar.h>
#include "include/unzip.h"
#include "include/zip.h"
#include <direct.h>

// 15 week template

using namespace std;

void extractAudioFiles(string dir);
int zipPPTs(string dir);
int unzipPPTs(string dir);
int createAudioDir(string dir);
int moveAudio(string dir, string folderName);

int main(int argc, char** argv)
{
	extractAudioFiles(string("C:\\Users\\Anthony\\Desktop\\CIT Project\\Test Folder\\")); //Replace with directory

	system("pause");
	return 0;
}

void extractAudioFiles(string dir)
{
	string pptDir;

	for (int i = 1; i < 17; i++)
	{
		if(i < 10)
		{
			pptDir = dir + "WEEK_0" + to_string(i) + "\\";
		}
		else
		{
			pptDir = dir + "WEEK_" + to_string(i) + "\\";
		}

		cout << "Current Directory: " + pptDir + "\n\n";

		if (createAudioDir(pptDir) == 0)
		{
			cout << "Audio_Conversion directory created successfully.\n\n";
		}

		zipPPTs(pptDir);
		pptDir += "Audio_Conversion\\";
		unzipPPTs(pptDir);
		cout << "----------------------------------------------------------------------------\n";
	}
}

int zipPPTs(string dir)
{
	DIR *dp;
	struct dirent *dirp;
	string oldFileName = "";

	if ((dp = opendir(dir.c_str())) == NULL)
	{
		cout << "Error(" << errno << ") opening " << dir << endl;
		return errno;
	}

	cout << "Powerpoint files found: \n";

	string folderDir;
	string name;
	string filePath;

	while ((dirp = readdir(dp)) != NULL)
	{
		if (string(dirp->d_name).find(".pptx") != std::string::npos && string(dirp->d_name).find(".zip") == std::string::npos)
		{
			oldFileName = (string(dir) + dirp->d_name); //Appends the directory to the front of the file name
			string newFileName = string(dir) + "Audio_Conversion\\" + dirp->d_name + ".zip";
			int result = rename(oldFileName.c_str(), newFileName.c_str());
			if(result != 0)
			{
				perror("Error converting file to zip");
			}
			else
			{
				cout << dirp->d_name << " changed to " << dirp->d_name << ".zip" << "\n";
			}
		}
		/*else if (string(dirp->d_name).find(".ppt") != std::string::npos && string(dirp->d_name).find(".zip") == std::string::npos)
		{
			name = (string)dirp->d_name;
			filePath = (string(dir) + name);
			string filePath2 = filePath + ".zip";
			HZIP hz;
			hz = CreateZip(_T(filePath2.c_str()), 0);
			ZipAdd(hz, name.c_str(), filePath.c_str());
			CloseZip(hz);
			cout << dirp->d_name << " changed to " << dirp->d_name << ".zip" << "\n";
		}*/
	}
	cout << "\n";
	closedir(dp);
	return 0;
}

int unzipPPTs(string dir)
{
	DIR *dp;
	struct dirent *dirp;
	string oldFileName = "";
	OFSTRUCT of;

	if ((dp = opendir(dir.c_str())) == NULL)
	{
		cout << "Error(" << errno << ") opening " << dir << endl;
		return errno;
	}

	string folderDir;
	string name;
	string filePath;

	cout << "Zip files extracted: \n";

	while ((dirp = readdir(dp)) != NULL)
	{
		name = (string)dirp->d_name;
		if (name.find(".zip") != std::string::npos)
		{
			filePath = (string(dir) + name);
			folderDir = dir + (name.substr(0,(name.length()-9))); //Remove .pptx.zip from file name
			if(_mkdir(folderDir.c_str()) == 0) //Create a new folder with file name to extract the zip file
			{
				HZIP hz;
				hz = OpenZip(_T(filePath.c_str()), 0);
				SetUnzipBaseDir(hz, _T(folderDir.c_str()));
				ZIPENTRY ze; GetZipItem(hz, -1, &ze); int numitems = ze.index;
				for (int zi = 0; zi<numitems; zi++)
				{
					GetZipItem(hz, zi, &ze);
					UnzipItem(hz, zi, ze.name);
				}
				CloseZip(hz);
				cout << dirp->d_name <<" extracted to " << folderDir << "\n";
			}
			string folderName = (string)dirp->d_name;
			moveAudio(folderDir, folderName.substr(0, folderName.length() - 9)); // 9 if pptx, 8 if ppt
		}
	}
	cout << "\n";
	closedir(dp);
	return 0;
}

int createAudioDir(string dir)
{
	dir += "Audio_Conversion";
	return _mkdir(dir.c_str());
}

int moveAudio(string dir, string folderName)
{
	DIR *dp;
	struct dirent *dirp;

	string newDir = dir + "\\ppt\\media\\";

	if ((dp = opendir(newDir.c_str())) == NULL)
	{
		cout << "Error(" << errno << ") opening " << dir << endl;
		return errno;
	}

	string name;
	string filePath;
	string newPath;

	cout<< "\n\tAudio files extracted: \n";

	while ((dirp = readdir(dp)) != NULL)
	{
		name = (string)dirp->d_name;
		if (name.find(".wav") != std::string::npos || name.find(".mp3") != std::string::npos || name.find(".mp4") != std::string::npos)
		{
			filePath = (string(newDir) + name);
			newPath = dir.substr(0, dir.length() - folderName.length()) + "\\" + folderName +  name;
			rename(filePath.c_str(), newPath.c_str());
			cout << "\t" << dirp->d_name << "\n";
		}
	}
	cout << "\n";
}