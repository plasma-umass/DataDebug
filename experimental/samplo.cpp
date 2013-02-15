// C++11
// clang++ -std=c++11 samplo.cpp

#include <algorithm>
#include <vector>
#include <iostream>
#include <string>
using namespace std;


#include <assert.h>
#include <math.h>
#include <stdlib.h>

/// @brief Generate a bootstrapped sample from the input distribution.
template <class TYPE>
void bootstrap (const vector<TYPE>& in,
		vector<TYPE>& out)
{
  assert (in.size() <= out.size());
  const int N = in.size();
  for (auto& x : out) {
    x = in[lrand48() % N];
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

int main()
{
  const auto NELEMENTS = 30;
  const auto NBOOTSTRAPS = 200;
  const auto ALPHA = 0.05; // = (1-alpha) confidence interval

  // Seed the random number generator.
  srand48 (time(NULL));

  vector<vectorType> a, b, impacts;

  a.resize (NELEMENTS);
  b.resize (NELEMENTS);
  impacts.resize (NELEMENTS);

  const float lambda = 0.01;
  // Generate a random vector.
  for (auto &x : a) {
    x = -log(drand48())/lambda;
    //    cout << "x = " << x << endl;
    //    x = (lrand48() % 1000) + 1;
  }
  // Add an anomalous value.
  a[8] = 1500;

  // Now we build the impact vector by bootstrapping.
 
  // For each index...
  for (long k = 0; k < NELEMENTS; k++) {
    long s = 0;
    // ...do a bunch of bootstraps, excluding that index, and add the results.
    for (long i = 0; i < NBOOTSTRAPS; i++) {
      exclusiveBootstrap(k, a, b);
      s += poly(b);
      //      cout << sum<vectorType>(NELEMENTS, b) << endl;
    }
    // The impact is the AVERAGE result over the bootstrapped samples.
    impacts[k] = s / NBOOTSTRAPS;
    cout << impacts[k] << endl;
  }

  auto sd   = stddev (impacts);
  auto mean = average (impacts);

  // Sort and then compute the confidence interval by
  // choosing the appropriate points in the vector.
  vector<vectorType> bt;
  bt.resize (NELEMENTS);
  unsigned int index = 0;
  for (auto& x : impacts) {
    bt[index++] = x;
  }
  sort (bt.begin(), bt.end());

  // Build a histogram.
  // This should be replaced by a hashmap or something.
  vector<int> freqCount;
  int length = bt[NELEMENTS-1]-bt[0]+1;
  freqCount.resize(length);
  for (int i = 0; i < NELEMENTS; i++) {
    freqCount[bt[i]]++;
  }

  cout << "# impact of injected anomaly = " << impacts[8] << endl;
  cout << "#  stddevs = " << ((float) impacts[8] - (float) mean) / (float) sd << endl;
  cout << "# mean impact = " << mean << endl;
  cout << "# stddev impact = " << sd << endl;
  int first;
  int last;

  cout << "# alpha = " << ALPHA << endl;
  cout << "# alpha/2 = " << ALPHA/2.0 << endl;
  cout << "# alpha/2 * newelements = " << (ALPHA/2.0 * (float) NELEMENTS) << endl;
  int tailcount = floor((ALPHA/2.0) * (float) NELEMENTS);
  cout << "# tailcount = " << tailcount << endl;

  {
    int count = 0;
    int index = 0;
    while (count + freqCount[bt[index]] <= tailcount) {
      count += freqCount[bt[index]];
      index++;
    }
    first = index;
  }

  {
    int count = 0;
    int index = NELEMENTS-1;
    while (count + freqCount[bt[index]] <= tailcount) {
      count += freqCount[bt[index]];
      index--;
    }
    last = index;
  }

  cout << "# " << 100.0 * (1.0-ALPHA) << "% confidence interval = [" << bt[first] << ", " << bt[last] << "]" << endl;

  // Now look for impacts that are outside the confidence interval.
  for (int i = 0; i < NELEMENTS; i++) {
    //    if (fabs((float) impacts[i] - (float) mean) >= 2 * sd) {
    if ((impacts[i] < bt[first]) ||
	(impacts[i] > bt[last])) {
      cout << "# item " << i << " appears anomalous: value = " << a[i] << ", impact = " << impacts[i] << ", stddevs = " << ((float) impacts[i] - (float) mean) / (float) sd << endl;
    }
  }


  return 0;
}
