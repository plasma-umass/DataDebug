// C++11
// clang++ -std=c++11 samplo.cpp

#include <vector>
#include <iostream>
#include <string>
using namespace std;


#include <assert.h>
#include <math.h>
#include <stdlib.h>

/// @brief Generate a bootstrapped sample from the input distribution.
template <class TYPE>
void bootstrap (int excludeIndex,
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
TYPE stddev (const vector<TYPE>& in) {
  TYPE avg = average (in);
  TYPE s = 0;
  for (auto const& x : in) {
    TYPE v = x - avg;
    s += (v * v);
  }
  return sqrt(s / (in.size()-1));
}

template <class TYPE>
TYPE average (const vector<TYPE>& in) {
  return sum (in) / in.size();
}

/*
 * A function, used for testing.
 *
 */

template <class TYPE>
TYPE poly (const vector<TYPE>& in) {
  TYPE s = 0;
  for (auto const& x : in) {
    s += (x > 700) ? 1 : 0;
  }
  return s; // pow(s, 0.5);
  //  return fabs(cos(s)) * 1000.0;
}


int main()
{
  const int NELEMENTS = 100;
  const int NBOOTSTRAPS = 1000;
  typedef unsigned long long vectorType;

  // Seed the random number generator.
  srand48 (time(NULL));

  vector<vectorType> a, b, impacts;

  a.resize (NELEMENTS);
  b.resize (NELEMENTS);
  impacts.resize (NELEMENTS);

  // Generate a random vector.
  for (auto &x : a) {
    x = (lrand48() % 1000) + 1;
  }
  // Add an anomalous value.
  a[8] = 2000;

  // Perform the computation.
  vectorType sumval = poly(a);

  // Now we build the impact vector by bootstrapping.

  // For each index...
  for (long k = 0; k < NELEMENTS; k++) {
    long s = 0;
    // ...do a bunch of bootstraps, excluding that index, and add the results.
    for (long i = 0; i < NBOOTSTRAPS; i++) {
      bootstrap(k, a, b);
      s += poly(b);
      //      cout << sum<vectorType>(NELEMENTS, b) << endl;
    }
    // The impact is the AVERAGE result over the bootstrapped samples.
    impacts[k] = s / NBOOTSTRAPS;
    cout << impacts[k] << endl;
  }
  return 0;
}
