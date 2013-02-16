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

const auto NELEMENTS = 30;
const auto NBOOTSTRAPS = 500;

// = (1-alpha) confidence interval
//  const auto ALPHA = 0.05; // 95% = 2 std devs
const auto ALPHA = 0.003; // 99.7% = 3 std devs

#include "realrandomvalue.h"
#include "mwc.h"

static MWC mwc (RealRandomValue::value(), RealRandomValue::value());

unsigned int rng() {
  return mwc.next();
};

/// @brief Fisher-Yates in-place shuffle.
template <class TYPE>
void shuffle (vector<TYPE>& vec)
{
  for (auto i = vec.size()-1; i != 0; i--) {
    auto j = rng() % i;
    swap (vec[i], vec[j]);
  }
}

/// @brief Fisher-Yates shuffle.
template <class TYPE>
void shuffle (const vector<TYPE>& in,
	      vector<TYPE>& out)
{
  out[0] = in[0];
  for (auto i = 1; i < in.size(); i++) {
    auto j = rng() % i;
    out[i] = out[j];
    out[j] = in[i];
  }
}


/// @brief Generate a bootstrapped sample from the input distribution.
template <class TYPE>
void bootstrap (const vector<TYPE>& in,
		vector<TYPE>& out)
{
  assert (in.size() <= out.size());
  const int N = in.size();
  for (auto& x : out) {
    x = in[rng() % N];
  }
}


/// @brief Generate a bootstrapped sample from the input distribution,
/// excluding one element.
template <class TYPE>
void exclusiveBootstrap (int excludeIndex,
			 const vector<TYPE>& in,
			 vector<TYPE>& out)
{
  assert (in.size() <= out.size());
  const int N = in.size();
  for (int i = 0; i < N; i++) {
    // Repeatedly pick an index at random to copy into the out array
    // (in other words, this is sampling WITH replacement).  If we hit
    // "excludeIndex", try again. Since this is unlikely to happen
    // frequently (on average, only once), it doesn't make much sense
    // to optimize.
    int index;
    index = excludeIndex;
    while (index == excludeIndex) {
      index = lrand48() % N;
    }
    out[i] = in[index];
  }
}

/*
 * Some basic stats functions over vectors.
 *
 */

template <class TYPE>
TYPE sum (const vector<TYPE>& in) {
  TYPE s = 0;
  for (auto const& x : in) {
    s += x;
  }
  return s;
}

template <class TYPE>
TYPE max (const vector<TYPE>& in) {
  TYPE m = in[0];
  for (auto& x : in) {
    if (x > m) {
      m = x;
    }
  }
  return m;
}

template <class TYPE>
TYPE average (const vector<TYPE>& in) {
  return sum (in) / in.size();
}

template <class TYPE>
TYPE stddev (const vector<TYPE>& in) {
  TYPE avg = average (in);
  TYPE s = 0;
  for (auto const& x : in) {
    TYPE v = x - avg;
    s += (v * v);
  }
  return sqrt(s / (in.size()-1));
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

typedef unsigned long long vectorType;

template <class TYPE>
int comparator (const void * a, const void * b) {
  auto ula = *((TYPE *) a);
  auto ulb = *((TYPE *) b);
  if (a < b) {
    return -1;
  }
  if (a == b) {
    return 0;
  }
  return 1;
}

template <class TYPE>
void computeOneBootstrap (const vector<TYPE>& mixsource,
			  int M,
			  int N,
			  float& bootstrap)
{
  vector<TYPE> mix;
  mix.resize (M + N);
  
  // Shuffle mix -- this will let us perform sampling without
  // replacement efficiently.
  shuffle (mixsource, mix);
  // Now we take the first M as "f", and the next N as "g".
  // Compute their means and save them in bootstrap.
  float sum1 = 0;
  for (auto i = 0; i < M; i++) {
    sum1 += mix[i];
  }
  float sum2 = 0;
  for (auto i = 0; i < N; i++) {
    sum2 += mix[M+i];
  }
  bootstrap = fabs((sum1/M) - (sum2/N));
}


template <class TYPE>
bool significantDifference (const float significanceLevel,
			    const vector<TYPE>& f,
			    const vector<TYPE>& g,
			    unsigned long NumBootstraps = 2000)
{
  assert (significanceLevel > 0.0);
  assert (significanceLevel < 1.0);

  // Compute the original mean.
  auto M = f.size();
  auto N = g.size();
  auto originalMean = fabs((float) average (f) - (float) average (g));

  // Combine both vectors.
  vector<TYPE> combined;
  combined.resize (M + N);
  int index = 0;
  for (auto const& x : f) {
    combined[index++] = x;
  }
  for (auto const& x : g) {
    combined[index++] = x;
  }

  // Build up the bootstrap of averages.
  vector<float> bootstrap;
  bootstrap.resize (NumBootstraps);

  for (auto i = 0; i < NumBootstraps; i++) {
    computeOneBootstrap(combined, M, N, bootstrap[i]);
  }
      
  // Now check to see whether the original mean is outside the
  // confidence interval.
  sort (bootstrap.begin(), bootstrap.end());
  
  int leftInterval = trunc(significanceLevel/2.0 * NumBootstraps);
  int rightInterval = trunc((1.0 - significanceLevel/2.0) * NumBootstraps);

  bool isOutside = ((originalMean < bootstrap[leftInterval]) ||
		    (originalMean > bootstrap[rightInterval]));
  return isOutside;
}


template <class TYPE>
bool significant (const int k,
		  const vector<TYPE>& original,
		  const vector<TYPE>& bootOriginal,
		  bool& result)
{
  vector<vectorType> b;
  b.resize (NELEMENTS);
  vector<vectorType> bootWithout;
  bootWithout.resize (NBOOTSTRAPS);
  for (long i = 0; i < NBOOTSTRAPS; i++) {
    exclusiveBootstrap(k, original, b);
    bootWithout[i] = poly(b);
  }
  result = significantDifference (ALPHA, bootOriginal, bootWithout, 10000);
  return result;
}


int main()
{
  // Seed the random number generator.
  srand48 (time(NULL));

  vector<vectorType> original;

  original.resize (NELEMENTS);

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
  
  // Bootstrap from the original sample.
  vector<vectorType> bootOriginal;
  bootOriginal.resize (NBOOTSTRAPS);
  vector<vectorType> b;
  b.resize (NELEMENTS);
  for (int i = 0; i < NBOOTSTRAPS; i++) {
    // Create a new bootstrap into b.
    bootstrap (original, b);
    // Compute the function and save it.
    bootOriginal[i] = poly (b);
  }

  thread * t = new thread[NELEMENTS];
  bool * sig = new bool[NELEMENTS];

  // For each index, check to see whether the distribution without it
  // is significantly different from the distribution with it (the
  // original).
  for (auto k = 0; k < NELEMENTS; k++) {
    significant (k, original, bootOriginal, sig[k]);
    // t[k] = thread (significant<vectorType>, k, a, boot, ref(sig[k]));
  }

  for (long k = 0; k < NELEMENTS; k++) {
    //    t[k].join();
    if (sig[k]) {
      cout << "element " << k << " (" << original[k] << ") significant." << endl;
    }
  }
  return 0;

}
