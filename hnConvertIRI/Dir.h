#pragma once
#include<iostream>
#include<vector>
using namespace std;
class Dir
{
public:
	Dir(void);
	~Dir(void);

	static bool Exist(string fullPath);
	static vector<string> GetDirectories(string fullPath);
};

