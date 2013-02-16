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
bool significantDifference (float significanceLevel,
			    const vector<TYPE>& f,
			    const vector<TYPE>& g,
			    unsigned long NumBootstraps = 2000)
{
  assert (significanceLevel > 0.0);
  assert (significanceLevel < 1.0);
  // Save the original mean.
  auto M = f.size();
  auto N = g.size();
  auto originalMean = fabs((float) average (f) - (float) average (g));

  // Combine both vectors into mix.
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
      
  // Now check to see whether the original mean is in the right
  // confidence interval.
  sort (bootstrap.begin(), bootstrap.end());
  
  int leftInterval = trunc(significanceLevel/2.0 * NumBootstraps);
  int rightInterval = trunc((1.0 - significanceLevel/2.0) * NumBootstraps);
  
  if ((originalMean < bootstrap[leftInterval]) ||
      (originalMean > bootstrap[rightInterval])) {
    return true;
  } else {
    return false;
  }
}


int main()
{
  const auto NELEMENTS = 30;
  const auto NBOOTSTRAPS = 1000;

  // = (1-alpha) confidence interval
  //  const auto ALPHA = 0.05; // 95% = 2 std devs
  const auto ALPHA = 0.003; // 99.7% = 3 std devs

  // Seed the random number generator.
  srand48 (time(NULL));

  vector<vectorType> a, b, impacts;

  a.resize (NELEMENTS);
  b.resize (NELEMENTS);
  impacts.resize (NELEMENTS);

  const float lambda = 0.01;

  // Generate a random vector.
  for (auto &x : a) {
    // Exponential distribution.
    x = -log(drand48())/lambda;
    cout << "# value = " << x << endl;
    //    x = (lrand48() % 750) + 1;
  }

  // Add an anomalous value.
  a[8] = 1000;
  
  // Bootstrap of original sample.
  vector<vectorType> boot;
  boot.resize (NBOOTSTRAPS);
  for (int i = 0; i < NBOOTSTRAPS; i++) {
    bootstrap (a, b);
    boot[i] = poly (b);
  }

  vector<vectorType> bootOne;
  bootOne.resize (NBOOTSTRAPS);

  // For each index...
  for (long k = 0; k < NELEMENTS; k++) {
    // ...do a bunch of bootstraps, excluding that index, and add the results.
    for (long i = 0; i < NBOOTSTRAPS; i++) {
      exclusiveBootstrap(k, a, b);
      bootOne[i] = poly(b);
    }
    cout << k << flush;
    if (significantDifference (0.003, boot, bootOne, 10000)) {
      cout << " SIGNIFICANT: value = " << a[k];
    }
    cout << endl;
  }
  return 0;

}
