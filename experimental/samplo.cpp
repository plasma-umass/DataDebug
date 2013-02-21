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

template <class TYPE>
void computeOneBootstrap (const vector<TYPE>& mixsource,
			  long M,
			  long N,
			  float& bootstrap)
{
  vector<TYPE> mix;
  mix.resize (M + N);
  
  // Shuffle mix -- this will let us perform sampling without
  // replacement efficiently.
  fyshuffle::transform (mixsource, mix);

  // Now we take the first M as "f", and the next N as "g".
  // Save the absolute difference of their means.
  float sum1 = 0;
  for (auto i = 0; i < M; i++) {
    sum1 += mix[i];
  }
  float sum2 = 0;
  for (auto i = 0; i < N; i++) {
    sum2 += mix[M+i];
  }
  bootstrap = (sum1/(float) M) - (sum2/(float) N);

  //  cout << "# boot avg = " << sum1/M << ", " << sum2 / N << endl << ": diff = " << bootstrap << endl;
}


template <class TYPE>
bool significantDifference (const float significanceLevel,
			    const vector<TYPE>& f,
			    const vector<TYPE>& g,
			    unsigned long NumBootstraps = NBOOTSTRAPS)
{
  assert (significanceLevel > 0.0);
  assert (significanceLevel < 1.0);

  // Compute the original difference in means.
  auto M = f.size();
  auto N = g.size();
  auto originalMeanDiff = (float) average (f) - (float) average (g);

  // Combine both vectors.
  vector<TYPE> combined;
  combined.resize (M + N);
  int index = 0;
  for (auto const& x : f) {
    combined[index++] = x;
  }
  assert (index == M);
  for (auto const& x : g) {
    combined[index++] = x;
  }
  assert (index == M + N);

  // Build up the bootstrap of averages.
  vector<float> bootstrap;
  bootstrap.resize (NumBootstraps);

  for (auto i = 0; i < NumBootstraps; i++) {
    computeOneBootstrap (combined, M, N, bootstrap[i]);
    // cout << "# avg bootstrap" << endl;
    // cout << bootstrap[i] << endl;
  }
      
  // Now check to see whether the original mean is outside the
  // confidence interval.
  sort (bootstrap.begin(), bootstrap.end());

  // Find the left and right intervals.
  int leftInterval = floor(significanceLevel / 2.0 * NumBootstraps);
  int rightInterval = ceil((1.0 - significanceLevel / 2.0) * NumBootstraps);

  cout << "# originalMeanDiff = " << originalMeanDiff << endl;
  cout << "# interval = ["
       << bootstrap[leftInterval] << ","
       << bootstrap[rightInterval] << "]" << endl;

  bool isOutside = ((originalMeanDiff < bootstrap[leftInterval]) ||
  		    (originalMeanDiff > bootstrap[rightInterval]));

  return isOutside;
}


template <class TYPE>
bool significant (const int k,
		  const vector<TYPE>& original,
		  const vector<TYPE>& bootOriginal,
		  bool& result)
{
  vector<TYPE> b;
  b.resize (NELEMENTS);

  // Build a bootstrap distribution WITHOUT index k.
  vector<TYPE> bootWithout;
  bootWithout.resize (NBOOTSTRAPS);
  for (long i = 0; i < NBOOTSTRAPS; i++) {
    bootstrap::exclusive (k, original, b);
    bootWithout[i] = poly(b) / (float) NELEMENTS;
    //    cout << "# boot without" << endl;
    //    cout << bootWithout[i] << endl;
  }
  // Now check to see if there's a significant difference in the
  // distribution means.

  assert (bootOriginal.size() == NBOOTSTRAPS);
  assert (bootWithout.size() == NBOOTSTRAPS);
  result = significantDifference (ALPHA, bootOriginal, bootWithout);


  cout << "# significant difference at " << (1.0-ALPHA) << " level? ";
  if (result) { cout << "YES"; } else { cout << "NO"; }
  cout << endl;
  cout << "# results for index " << k << endl;
  cout << "# avg with = " << average (bootOriginal) << endl;
  cout << "# avg without = " << average (bootWithout) << endl;
  return result;
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

  sort (bootOriginal.begin(), bootOriginal.end());

  //  cout << "# bootOriginal average = " << average (bootOriginal) << endl;
  //  cout << "# bootOriginal SD = " << stddev (bootOriginal) << endl;

  bool * sig = new bool[NELEMENTS];

  // For each index, check to see whether the distribution without it
  // is significantly different from the distribution with it (the
  // original).
  vector<float> bootWithout[NELEMENTS];

  auto bootOrigAvg = average (bootOriginal);

  // Find the left and right intervals.
  auto leftInterval = floor(ALPHA / 2.0 * NBOOTSTRAPS);
  auto rightInterval = ceil((1.0 - ALPHA / 2.0) * NBOOTSTRAPS);

  for (auto k = 0; k < NELEMENTS; k++) {

    bootWithout[k].resize (NBOOTSTRAPS);

    // Build a bootstrap distribution WITHOUT index k.
    for (long i = 0; i < NBOOTSTRAPS; i++) {
      bootstrap::exclusive (k, original, b);
      bootWithout[k][i] = poly(b) / (float) NELEMENTS;
    }

    sort (bootWithout[k].begin(), bootWithout[k].end());

    // KS test.

    auto max = -1.0;
    for (auto i = 0; i < NBOOTSTRAPS; i++) {
      auto val = fabs(bootOriginal[i]- bootWithout[k][i]);
      if (val > max) {
	max = val;
      }
    }
    // c(0.001) = 1.95
    // Reject the null hypothesis if KS > c(alpha) * critical value.
    auto criticalValue = 2.0 * sqrt(((NBOOTSTRAPS*NBOOTSTRAPS)/(2.0 * NBOOTSTRAPS)));
    //    auto criticalValue = 2.0 * sqrt(((NBOOTSTRAPS*NBOOTSTRAPS)/(2.0 * NBOOTSTRAPS)));
    auto KS = sqrt(((NBOOTSTRAPS*NBOOTSTRAPS)/(2.0 * NBOOTSTRAPS)) * max);
    if (KS > criticalValue) {
      cout << "#element " << k << " is significantly different." << endl;
    }
  }

  return 0;

}
