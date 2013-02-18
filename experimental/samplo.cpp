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

const auto NELEMENTS = 50;
const auto NBOOTSTRAPS = 1000;

// = (1-alpha) confidence interval
const auto ALPHA = 0.05; // 95% = 2 std devs
// const auto ALPHA = 0.003; // 99.7% = 3 std devs
// const auto ALPHA = 0.001;

#include "fyshuffle.h"
#include "stats.h"

using namespace fyshuffle;
using namespace stats;

/// @brief Generate a bootstrapped sample from the input distribution.
template <class TYPE>
void bootstrap (const vector<TYPE>& in,
		vector<TYPE>& out)
{
  assert (in.size() == out.size());
  const auto N = in.size();
  for (auto& x : out) {
    x = in[lrand48() % N];
  }
}


/// @brief Generate a bootstrapped sample from the input distribution,
/// excluding one element.
template <class TYPE>
void exclusiveBootstrap (unsigned long excludeIndex,
			 const vector<TYPE>& in,
			 vector<TYPE>& out)
{
  assert (in.size() == out.size());
  const auto N = in.size();
  for (auto i = 0; i < N; i++) {
    // Repeatedly pick an index at random to copy into the out array
    // (in other words, this is sampling WITH replacement).  If we hit
    // "excludeIndex", try again. Since this is unlikely to happen
    // frequently (on average, only once), it doesn't make much sense
    // to optimize.
    auto index = excludeIndex;
    while (index == excludeIndex) {
      index = lrand48() % N;
    }
    out[i] = in[index];
    //    cout << "# exclusive boot " << excludeIndex << " - " << out[i] << endl;
  }
}

/*
 * A function, used for testing.
 *
 */

template <class TYPE>
TYPE poly (const vector<TYPE>& in) {
  TYPE s = 0;
  for (auto const& x : in) {
    //    s += (x > 700) ? 1 : 0;
    s += x; // x * x;
  }
  return s; 
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
  bootstrap = fabs((sum1/(float) M) - (sum2/(float) N));

  //  cout << "# boot avg = " << sum1/M << ", " << sum2 / N << endl << ": diff = " << bootstrap << endl;
}


template <class TYPE>
bool significantDifference (const float significanceLevel,
			    const vector<TYPE>& f,
			    const vector<TYPE>& g,
			    unsigned long NumBootstraps = 10000)
{
  assert (significanceLevel > 0.0);
  assert (significanceLevel < 1.0);

  // Compute the original difference in means.
  auto M = f.size();
  auto N = g.size();
  auto originalMeanDiff = fabs((float) average (f) - (float) average (g));

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

  float avgBoot = average (bootstrap);

  // Find the left and right intervals.
  int leftInterval = floor(significanceLevel / 2.0 * NumBootstraps);
  int rightInterval = ceil((1.0 - significanceLevel / 2.0) * NumBootstraps);

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
    exclusiveBootstrap(k, original, b);
    bootWithout[i] = poly(b); // / (float) NELEMENTS;
    //    cout << "# boot without" << endl;
    //    cout << bootWithout[i] << endl;
  }
  // Now check to see if there's a significant difference in the
  // distribution means.

  assert (bootOriginal.size() == NBOOTSTRAPS);
  assert (bootWithout.size() == NBOOTSTRAPS);
  result = significantDifference (ALPHA, bootOriginal, bootWithout);

  //  if (k == 0) {
  //    for (auto x : bootWithout) {
  //      cout << x << endl;
  //    }
  //  }

  //  if (result) {
  cout << "# significant difference at " << (1.0-ALPHA) << " level? ";
  if (result) { cout << "YES"; } else { cout << "NO"; }
  cout << endl;
  cout << "# results for index " << k << endl;
  cout << "# avg with = " << average (bootOriginal) << endl;
  cout << "# avg without = " << average (bootWithout) << endl;
  //  cout << "# stddev original = " << stddev (bootOriginal) << endl;
  //  }
  return result;
}

typedef float vectorType;

int main()
{
  // Seed the random number generator.
  srand48 (0); // time(NULL));

#if 0
  // Testing shuffle.
  vector<vectorType> q, r;
  q.resize(5);
  r.resize(5);
  q[0] = 1;
  q[1] = 2;
  q[2] = 3;
  q[3] = 4;
  q[4] = 5;
  shuffle::transform (q, r);
  for (auto x : r) {
    cout << "r = " << x << endl;
  }
  return 0;
#endif

  vector<vectorType> original;

  original.resize (NELEMENTS);

  // Generate a random vector.
  for (auto &x : original) {
    // Uniform distribution.
    x = lrand48() % 9 + 1;
  }

  original[2] = 40;
  
  
#if 0
  // Generate a random vector.
  const float lambda = 0.01;
  for (auto &x : original) {
    // Exponential distribution.
    x = -log(drand48())/lambda;
    cout << "# value = " << x << endl;
    //    x = (lrand48() % 750) + 1;
  }

  // Add an anomalous value.
  original[8] = 1000;
#endif
  
 
  for (auto const& x : original) {
    cout << "# value = " << x << endl;    
  }
  
  // Bootstrap from the original sample.
  vector<vectorType> bootOriginal;
  bootOriginal.resize (NBOOTSTRAPS);
  vector<vectorType> b;
  b.resize (NELEMENTS);
  for (int i = 0; i < NBOOTSTRAPS; i++) {
    // Create a new bootstrap into b.
    bootstrap (original, b);
    // Compute the function and save it.
    bootOriginal[i] = poly (b); //  / (float) NELEMENTS;
    //    cout << bootOriginal[i] << " # " << __FILE__ << ":" << __LINE__ << endl;
  }

  sort (bootOriginal.begin(), bootOriginal.end());

  //  cout << "# bootOriginal average = " << average (bootOriginal) << endl;
  //  cout << "# bootOriginal SD = " << stddev (bootOriginal) << endl;

  bool * sig = new bool[NELEMENTS];

  // For each index, check to see whether the distribution without it
  // is significantly different from the distribution with it (the
  // original).
  vector<float> bootWithout;
  bootWithout.resize (NBOOTSTRAPS);

  auto bootOrigAvg = average (bootOriginal);

  // Find the left and right intervals.
  auto leftInterval = floor(ALPHA / 2.0 * NBOOTSTRAPS);
  auto rightInterval = ceil((1.0 - ALPHA / 2.0) * NBOOTSTRAPS);

  for (auto k = 0; k < NELEMENTS; k++) {

    // Build a bootstrap distribution WITHOUT index k.
    for (long i = 0; i < NBOOTSTRAPS; i++) {
      exclusiveBootstrap(k, original, b);
      bootWithout[i] = poly(b); //  / (float) NELEMENTS;
    }

    auto bootWoAvg = average (bootWithout);

    sort (bootWithout.begin(), bootWithout.end());

    cout << "# bootOriginal AVERAGE = " << average (bootOriginal) << endl;
    cout << "# bootWithout  AVERAGE = " << average (bootWithout) << endl;
    cout << "# bootOriginal[25] = "  << bootOriginal[25] << endl;
    cout << "# bootOriginal[975] = " << bootOriginal[975] << endl;
    cout << "# bootWithout[25] = "   << bootWithout[25] << endl;
    cout << "# bootWithout[975] = "  << bootWithout[975] << endl;

    if ((bootWoAvg < bootOriginal[leftInterval]) ||
	(bootWoAvg > bootOriginal[rightInterval]) ||
	(bootOrigAvg < bootWithout[leftInterval]) ||
	(bootOrigAvg > bootWithout[rightInterval])) {
      sig[k] = true;
    } else {
      sig[k] = false;
    }


    //    sig[k] = significantDifference (0.001, bootOriginal, bootWithout, 1000);
  }

  for (long k = 0; k < NELEMENTS; k++) {
    //    t[k].join();
    if (sig[k]) {
      cout << "# element " << k << " (" << original[k] << ") significant." << endl;
    }
  }
  return 0;

}
