#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "Test.h"

int main()
{
	// Initial string
	char* input = calloc(20, sizeof(char));
	strcpy(input, "AnaAreMere");
	printf("String: %s\n", input);

	// Reversed string
	mirror(input);
	printf("Reversed string: %s\n", input);
	free(input);

	return 0;
}