// C++11
// clang++ -std=c++11 -stdlib=libc++ samplo.cpp

#include <algorithm>
#include <vector>
#include <iostream>
#include <map>
#include <string>
#include <thread>
using namespace std;


#include <assert.h>
#include <math.h>
#include <stdlib.h>

const auto NELEMENTS = 100;
const auto NBOOTSTRAPS = 5000;

// = (1-alpha) confidence interval
// const auto ALPHA = 0.05; // 95% = 2 std devs
const auto ALPHA = 0.003; // 99.7% = 3 std devs
// const auto ALPHA = 0.001;

#include "fyshuffle.h"
#include "stats.h"
#include "bootstrap.h"

using namespace fyshuffle;
using namespace stats;


/*
 * A function, used for testing.
 *
 */

template <class TYPE>
TYPE poly (const vector<TYPE>& in) {
  auto s = 0.0;
  for (auto const& x : in) {
    s += (x > 7.0) ? (x * 1.0) : (x * 0.1);
    // s += cos(x * x); // x * x;
  }
  return s; // sqrt(s); 
  // return sqrt(s);
}

typedef float vectorType;

int main()
{
  // Seed the random number generator.
  srand48 (0); // time(NULL));

  vector<vectorType> original;
  original.resize (NELEMENTS);

#if 1

  // Generate a random vector.
  for (auto &x : original) {
    // Uniform distribution.
    x = lrand48() % 3 + 1;
  }

  // Add an anomalous value.
  original[2] = 180; // 640; // 64;
   
#else

  // Generate a random vector.
  const float lambda = 0.01;
  for (auto &x : original) {
    // Exponential distribution.
    x = -log(drand48())/lambda;
  }

  // Add an anomalous value.
  //  original[8] = 1000;
#endif
  
 
#if 1
  for (auto const& x : original) {
    cout << "# value = " << x << endl;    
  }
#endif
  
  // Bootstrap from the original sample.
  vector<vectorType> bootOriginal;
  bootOriginal.resize (NBOOTSTRAPS);
  vector<vectorType> b;
  b.resize (NELEMENTS);
  for (int i = 0; i < NBOOTSTRAPS; i++) {
    // Create a new bootstrap into b.
    bootstrap::complete (original, b);
    // Compute the function and save it.
    bootOriginal[i] = poly (b) / (float) NELEMENTS;
    cout << bootOriginal[i] << " # " << __FILE__ << ":" << __LINE__ << endl;
  }

  // For each index, check to see whether the distribution without it
  // is significantly different from the distribution with it (the
  // original).

  vector<float> bootWithout;
  bootWithout.resize (NBOOTSTRAPS);

  for (auto k = 0; k < NELEMENTS; k++) {

    // Build a bootstrap distribution WITHOUT index k.
    for (long i = 0; i < NBOOTSTRAPS; i++) {
      bootstrap::exclusive (k, original, b);
      bootWithout[i] = poly(b) / (float) NELEMENTS;
    }

    if (stats::kolmogorovSmirnoff (bootOriginal, bootWithout)) {
      cout << "#element " << k << " is significantly different." << endl;
    }
  }

  return 0;

}
