#ifndef __CONVERT_H_INCLUDED__
#define __CONVERT_H_INCLUDED__

// Temperature Conversion

// from Fahrenheit
double FahrenheitToCelsius ( double x ) { return 5.0 / 9.0 * ( x - 32.0 ); }
double FahrenheitToKelvin ( double x ) { return x + 459.67; }

// from Celsius
double CelsiusToFahrenheit ( double x ) { return 9.0 / 5.0 * ( x + 32 ); }
double CelsiusToKelvin ( double x ) { return x + 273.15; }

// from Kelvin
double KelvinToFahrenheit (double x ) { return x - 459.67; }  // don't believe this is correct ?
double KelvinToCelsius ( double x ) { return x - 273.15; }

// Measurement conversion

