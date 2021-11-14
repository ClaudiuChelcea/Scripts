// Calculate length without dirrect accessing with []
int my_strlen(const char *str)
{
	const char* start = str;
	int length = 0;

	while(*start)
	{
		length++;
		start++;
	}

	return length;
}

void mirror(char *s)
{
	int length = my_strlen(s);
	
	for(int i = 0; i < length / 2; ++i)
	{
		char tmp = *(s + i);
		*(s + i) = *(s + length - 1 - i);
		*(s + length-1-i) = tmp;
	}
}