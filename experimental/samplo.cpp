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

const auto NELEMENTS = 20;
const auto NBOOTSTRAPS = 2000;

// = (1-alpha) confidence interval
// const auto ALPHA = 0.05; // 95% = 2 std devs
//const auto ALPHA = 0.003; // 99.7% = 3 std devs
const auto ALPHA = 0.001;

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
    s += x;
    //    s += (x > 7.0) ? (x * 1.0) : (x * 0.0);
    // s += cos(x * x); // x * x;
  }
  return s; // sqrt(s); 
  // return sqrt(s);
}

typedef float vectorType;

int main()
{
  // Seed the random number generator.
  srand48 (time(NULL));

  vector<vectorType> original;
  original.resize (NELEMENTS);

#if 1

  // Generate a random vector.
  for (auto &x : original) {
    // Uniform distribution.
    x = lrand48() % 100 + 1;
  }

  // Add an anomalous value.
  original[2] = 180; // 4; // 180; // 640; // 64;
   
#else

  // Generate a random vector.
  const float lambda = 0.01;
  for (auto &x : original) {
    // Exponential distribution.
    x = -log(drand48())/lambda;
  }

  // Add an anomalous value.
  // original[8] = 1000;
#endif
  
 
#if 1
  {
    int count = 0;
    for (auto const& x : original) {
      cout << "(" << count << ") value = " << x << endl;    
      count++;
    }
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
    bootOriginal[i] = poly (b);
    //    cout << bootOriginal[i] << " # " << __FILE__ << ":" << __LINE__ << endl;
  }

#if 1
  vector<int> excludes[NELEMENTS];
  float boots[NBOOTSTRAPS];
  const auto N = NELEMENTS;
  auto overallSum = 0.0;

  // Build up the boots (bootstrap) array of values
  // from the original distribution.

  // At the same time, organize them so that each index of the
  // excludes array comprises all of the bootstrapped values that do
  // NOT contain a given indexed value.
  for (int i = 0; i < NBOOTSTRAPS; i++) {
    vector<float> out;
    out.resize (NELEMENTS);

    vector<bool> includedPosition;
    includedPosition.resize (NELEMENTS);
    bootstrap::completeTracked (original, out, includedPosition);
    auto result = poly (out);
    boots[i] = result;
    overallSum += result;
    
    // Check each included position index and update excludes
    // accordingly.
    for (auto k = 0; k < N; k++) {
      if (!includedPosition[k]) {
	excludes[k].push_back (i);
      }
    }

    // Sort the excludes array. We will use this in subsequent steps.
    for (auto k = 0; k < N; k++) {
      sort (excludes[k].begin(), excludes[k].end());
    }

  }
  auto overallAvg = overallSum / NBOOTSTRAPS;

  cout << "overall avg = " << overallAvg << endl;

  for (auto k = 0; k < N; k++) {
    // Compute the mean without this index.
    auto sum = 0.0;
    for (auto ind : excludes[k]) {
      sum += boots[ind];
    }
    auto avg = sum / (float) excludes[k].size();
    cout << "avg WITHOUT element " << k << "= " << avg << endl;

    // Now compute the distribution of values WITH this element...
    vector<float> distrib;
    auto excIndex = 0;
    for (auto i = 0; i < NBOOTSTRAPS; i++) {
      if ((excIndex < excludes[k].size()) && (i == excludes[k][excIndex])) {
	excIndex++;
      } else {
	distrib.push_back (boots[i]);
      }
    }
    sort (distrib.begin(), distrib.end());

    auto const sz = distrib.size();
    cout << "[" << distrib[0.025 * sz] << "," << distrib[0.975 * sz] << "]" << endl;

    if ((avg < distrib[0.025 * sz]) || (avg > distrib[0.975 * sz])) {
      cout << "SIGNIFICANT!!" << endl;
    }
  }

  // 
#endif

#if 0
  // For each index, check to see whether the distribution without it
  // is significantly different from the distribution with it (the
  // original).

  vector<float> bootWithout;
  bootWithout.resize (NBOOTSTRAPS);

  for (auto k = 0; k < NELEMENTS; k++) {

    // Build a bootstrap distribution WITHOUT index k.
    for (long i = 0; i < NBOOTSTRAPS; i++) {
      bootstrap::exclusive (k, original, b);
      bootWithout[i] = poly(b);
    }

    //    cout << "overlap fraction = " << overlapFraction (bootOriginal, bootWithout) << endl;

    
#if 0
    auto f = confidencePermutationTest (bootOriginal, bootWithout, 10000);
    cout << "f value = " << f << endl;
    if (f <= 0.0001) {
      cout << "#element " << k << " (" << original[k] << ") is significantly different per permutation test." << endl;
    }
#endif

#if 1
    if (stats::meanDistance (bootOriginal, bootWithout)) {
      cout << "#element " << k << " is significantly different per mean-distance test." << endl;
    }
#endif

#if 0
    if (stats::kolmogorovSmirnoff (bootOriginal, bootWithout)) {
      cout << "#element " << k << " is significantly different per KS test." << endl;
    }
#endif

#if 1
    if (stats::mannWhitney (bootOriginal, bootWithout, 0.0001)) {
      cout << "#element " << k << " is significantly different per Mann-Whitney test." << endl;
    }

#endif
  }
#endif

  return 0;

}
