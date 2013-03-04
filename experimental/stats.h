// -*- C++ -*-

#ifndef STATS_H_
#define STATS_H_

#include <assert.h>
#include <math.h>

#include <vector>
#include <algorithm>
using namespace std;

#include "bootstrap.h"

/*
 * Some basic stats functions over vectors.
 *
 */

namespace stats {

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
  float average (const vector<TYPE>& in) {
    return sum (in) / (float) in.size();
  }

  template <class TYPE>
  float stddev (const vector<TYPE>& in) {
    auto avg = average (in);
    float s = 0;
    for (auto const& x : in) {
      auto v = x - avg;
      s += (v * v);
    }
    return sqrt(s / (in.size()-1));
  }

  template <class TYPE>
  int rankCount (vector<TYPE>& a,
		 vector<TYPE>& b)
  {
    int count = 0;
    sort (a.begin(), a.end());
    sort (b.begin(), b.end());
    
    int aIndex = 0;
    int bIndex = 0;
    
    while ((aIndex < a.size()) && (bIndex < b.size())) {
      if (a[aIndex] < b[bIndex]) {
	count += b.size() - bIndex;
	aIndex++;
      } else {
	bIndex++;
      }
    }
    return count;
  }


  // Non-parametric form (via Monte Carlo) of Mann-Whitney test:
  // MC is to compute p-value.
  template <class TYPE>
  bool mannWhitney (vector<TYPE>& a,
		    vector<TYPE>& b,
		    float significanceLevel = 0.05)
  {
    assert (a.size() == b.size());

    sort (a.begin(), a.end());
    sort (b.begin(), b.end());

    auto N = a.size() + b.size();
    auto MaxRank = N * (N + 1) / 2.0;

    int count = rankCount<TYPE> (a, b);

    cout << "count = " << count << endl;

    vector<TYPE> combined;

    for (auto x : a) { combined.push_back (x); }
    for (auto x : b) { combined.push_back (x); }

    // Now see what the probability of this value is with Monte Carlo.
    int freq = 0;

    const int ITERATIONS = 5000;

    for (int it = 0; it < ITERATIONS; it++) {
      fyshuffle::inplace (combined);
      
      vector<TYPE> firstVector (combined.begin(), combined.begin() + a.size());
      vector<TYPE> secondVector (combined.begin() + a.size(), combined.end());
      int r = rankCount<TYPE>(firstVector, secondVector);
      if ((r < count) || (r > MaxRank-count)) {
	freq++;
      }
    }

    cout << "freq = " << freq << " out of " << ITERATIONS << endl;

    if ((float) freq / ITERATIONS < significanceLevel) {
      return true;
    } else {
      return false;
    }

    return false;

#if 0    
    sort (a.begin(), a.end());
    sort (b.begin(), b.end());

    vector<pair<TYPE,int>> combined;

    for (auto x : a) { combined.push_back (pair<TYPE,int>(x+drand48(),0)); }
    for (auto x : b) { combined.push_back (pair<TYPE,int>(x+drand48(),1)); }

    sort (combined.begin(), combined.end());

    // Compute ranks. For now, ignore ties.
    auto index = 1;
    auto aRanks = 0;
    for (auto x : combined) {
      if (x.second == 0) {
	aRanks += index;
      }
      index++;
    }
    
    cout << "RANKS = " << aRanks << endl;

    auto n1 = a.size();
    auto n2 = b.size();
    auto N  = n1 + n2;

    auto U = aRanks - (n1 * (n1+1))/2.0;
    cout << "aRanks = " << aRanks << endl;
    cout << "U = " << U << endl;
    auto s = sqrt ((n1*n2)*(n1+n2+1.0)/12.0);
    auto Z = fabs((U - (n1*n2)/2.0) / s);
    cout << "Z = " << Z << endl;
    return false;
#endif
  }

  void testMannWhitney() {
    vector<int> a = {0, 6, 7, 8, 9, 10};
    vector<int> b = {1, 2, 3, 4, 5, 11};
    cout << mannWhitney (a, b) << endl;
  }

  template <class TYPE>
  bool meanDistance (vector<TYPE>& a,
		     vector<TYPE>& b,
		     float significanceLevel = 0.05)
  {
    assert (a.size() == b.size());

    sort (a.begin(), a.end());
    sort (b.begin(), b.end());

    //    auto aAverage = a[a.size()/2]; // average (a);
    //    auto bAverage = b[b.size()/2]; // average (b);
    auto aAverage = average (a);
    auto bAverage = average (b);

    auto aLeftInterval  = floor(significanceLevel / 2.0 * a.size());
    auto aRightInterval = ceil((1.0 - significanceLevel / 2.0) * a.size());

    auto bLeftInterval  = floor(significanceLevel / 2.0 * b.size());
    auto bRightInterval = ceil((1.0 - significanceLevel / 2.0) * b.size());

    bool result;
    if ((aAverage < b[bLeftInterval]) ||
	(aAverage > b[bRightInterval]) ||
	(bAverage < a[aLeftInterval]) ||
	(bAverage > a[aRightInterval])) {
      result = true;
    } else {
      result = false;
    }
    return result;
  }

  template <class TYPE>
  float confidencePermutationTest (const vector<TYPE>& a,
				   const vector<TYPE>& b,
				   const int iterations = 1000)
  {
    auto originalMeanDiff = fabs((float) average (a) - (float) average (b));

    cout << "original mean diff = " << originalMeanDiff << endl;

    // Combine a and b into a vector called mix.
    vector<TYPE> mix;
    mix.resize (a.size() + b.size());

    {
      auto index = 0;
      for (auto const& x : a) {
	mix[index++] = x;
      }
      for (auto const& x : b) {
	mix[index++] = x;
      }
    }
    
    // Now repeatedly construct two different permutations of this mix,
    // and compute their averages. We count the number of times the original average
    // is smaller than the average of the permuted samples.
    auto count = 0;
    float minDiff = 1e99;
    float maxDiff = -1e99;
    for (auto i = 0; i < iterations; i++) {
      fyshuffle::inplace (mix);
      float s1 = 0.0, s2 = 0.0;
      for (auto j = 0; j < a.size(); j++) {
	s1 += mix[j];
      }
      for (auto j = a.size(); j < a.size() + b.size(); j++) {
	s2 += mix[j];
      }
      float currDiff = fabs(s1/a.size() - s2/b.size());
      if (minDiff > currDiff) { minDiff = currDiff; }
      if (maxDiff < currDiff) { maxDiff = currDiff; }
      if (originalMeanDiff <= currDiff) {
	count++;
      }
    }
    cout << "count = " << count << endl;
    cout << "minDiff = " << minDiff << endl;
    cout << "maxDiff = " << maxDiff << endl;
    return (float) count / (float) iterations;
  }

  template <class TYPE>
  float overlapFraction (vector<TYPE>& a,
			 vector<TYPE>& b)
  {
    sort (a.begin(), a.end());
    sort (b.begin(), b.end());

    auto counter = 0;

    vector<TYPE> vecs[2];
    vecs[0] = a;
    vecs[1] = b;
    float range[2][2] = {{ a[0], a[a.size()-1] },
			 { b[0], b[b.size()-1] }};

    for (int i = 0; i < 2; i++) {
      for (int index = 0; index < vecs[i].size(); index++) {
	if ((vecs[i][index] >= range[1-i][0]) &&
	    (vecs[i][index] <= range[1-i][1])) {
	  counter++;
	}
      }
    }

    return (float) counter / (float) (a.size() + b.size());
  }

  /// @brief returns true iff a and b are significantly different.
  /// i.e., it's safe to reject the null hypothesis that they are from
  /// the same distribution.
  template <class TYPE>
  bool kolmogorovSmirnoff (vector<TYPE>& a,
			   vector<TYPE>& b,
			   float significanceLevel = 0.001)
  {
    sort (a.begin(), a.end());
    sort (b.begin(), b.end());
    assert (a.size() == b.size());
    auto max = -1.0;
    for (auto i = 0; i < a.size(); i++) {
      auto val = fabs(a[i]- b[i]);
      if (val > max) {
	max = val;
      }
    }
    // c(0.001) = 1.95
    // Reject the null hypothesis if KS > c(alpha) * critical value.
    // NB: right now we just use a significance level of 0.001!
    auto criticalValue = 2.0 * sqrt(((a.size()*a.size())/(2.0 * a.size())));
    auto KS = sqrt(((a.size()*a.size())/(2.0 * a.size())) * max);
    return (KS > criticalValue);
  }


  template <class TYPE>
  void bs2 (const vector<TYPE>& original,
	    TYPE func (const vector<TYPE>&),
	    vector<bool>& significant,
	    const int nBootstraps = 2000,
	    const float significanceLevel = 0.05)
  {
    // All of the indices that do NOT contain the given element.
    vector<vector<int>> excludes;
    excludes.resize (original.size());

    TYPE boots[nBootstraps];
    const auto N = original.size();
    auto overallSum = 0.0;

    vector<pair<TYPE,vector<bool>>> bootstrap;
    bootstrap.resize (nBootstraps);

    // Build up the boots (bootstrap) array of values
    // from the original distribution.
    
    // At the same time, organize them so that each index of the
    // excludes array comprises all of the bootstrapped values that do
    // NOT contain a given indexed value.
    for (int i = 0; i < nBootstraps; i++) {
      vector<TYPE> out;
      out.resize (original.size());
      
      bootstrap[i].second.resize (original.size());

      bootstrap::completeTracked (original, out, bootstrap[i].second);
      auto result = func (out);
      bootstrap[i].first = result;
      overallSum += result;
    }

    sort (bootstrap.begin(), bootstrap.end());

    for (auto i = 0; i < nBootstraps; i++) {

      boots[i] = bootstrap[i].first;

      // Check each included position index and update excludes
      // accordingly.
      for (auto k = 0; k < N; k++) {
	if (!bootstrap[i].second[k]) {
	  excludes[k].push_back (i);
	}
      }
    }

    for (auto k = 0; k < N; k++) {

      vector<int> ind;
      ind.resize (nBootstraps);
      for (auto i = 0; i < nBootstraps; i++) {
	ind[i] = i;
      }

      vector<float> meanDiff;
      meanDiff.resize (1000);
      for (auto i = 0; i < 1000; i++) {
	// Repeatedly divy up all the indices into two sets,
	// and record the difference of the means.
	fyshuffle::inplace (ind);
	auto sum_n1 = 0;
	auto n1 = 0;
	for (n1 = 0; n1 < excludes[k].size(); n1++) {
	  sum_n1 += boots[ind[n1]];
	}

	auto sum_n2 = 0;
	auto n2 = 0;
	for (n2 = 0; n2 < nBootstraps - excludes[k].size(); n2++) {
	  sum_n2 += boots[ind[n2+excludes[k].size()]];
	}

	meanDiff[i] = fabs((sum_n1 / (float) n1) - (sum_n2 / (float) n2));
      }

      // Compute the mean without this index.
      auto sum = 0.0;
      for (auto ind : excludes[k]) {
	sum += boots[ind];
      }
      auto avgWithout = sum / (float) excludes[k].size();

      // Now the mean WITH this index.
      // Now compute the distribution of values WITH this element...
      auto excIndex = 0;
      sum = 0;
      for (auto i = 0; i < nBootstraps; i++) {
	if ((excIndex < excludes[k].size()) && (i == excludes[k][excIndex])) {
	  excIndex++;
	} else {
	  sum += boots[i];
	}
      }
      auto avgWith = sum / (float) (nBootstraps - excludes[k].size());
      auto avgDiff = fabs (avgWith - avgWithout);

      // Where does this fall in meanDiff?
      sort (meanDiff.begin(), meanDiff.end());

      auto const sz = 1000.0;
      int leftInterval  = floor (significanceLevel / 2.0 * sz);
      int rightInterval = ceil ((1.0 - significanceLevel / 2.0) * sz - 1);

      if ((avgDiff < meanDiff[leftInterval]) || (avgDiff > meanDiff[rightInterval])) {
	cout << "(" << k << ") avgDiff = " << avgDiff << ", interval = [" << meanDiff[leftInterval] << "," << meanDiff[rightInterval] << "]" << endl;

	significant[k] = true;
      } else {
	significant[k] = false;
      }
    }
  }


  template <class TYPE>
  void withAndWithoutYou (const vector<TYPE>& original,
			  TYPE func (const vector<TYPE>&),
			  vector<bool>& significant,
			  const int nBootstraps = 2000,
			  const float significanceLevel = 0.05)
  {
    // All of the indices that do NOT contain the given element.
    vector<vector<int>> excludes;
    excludes.resize (original.size());

    TYPE boots[nBootstraps];
    const auto N = original.size();
    auto overallSum = 0.0;

    vector<pair<TYPE,vector<bool>>> bootstrap;
    bootstrap.resize (nBootstraps);

    // Build up the boots (bootstrap) array of values
    // from the original distribution.
    
    // At the same time, organize them so that each index of the
    // excludes array comprises all of the bootstrapped values that do
    // NOT contain a given indexed value.
    for (int i = 0; i < nBootstraps; i++) {
      vector<TYPE> out;
      out.resize (original.size());
      
      bootstrap[i].second.resize (original.size());

      bootstrap::completeTracked (original, out, bootstrap[i].second);
      auto result = func (out);
      bootstrap[i].first = result;
      overallSum += result;
    }

    sort (bootstrap.begin(), bootstrap.end());

    for (auto i = 0; i < nBootstraps; i++) {

      boots[i] = bootstrap[i].first;

      // Check each included position index and update excludes
      // accordingly.
      for (auto k = 0; k < N; k++) {
	if (!bootstrap[i].second[k]) {
	  excludes[k].push_back (i);
	}
      }
    }

    for (auto k = 0; k < N; k++) {
      // Compute the mean without this index.
      auto sum = 0.0;
      for (auto ind : excludes[k]) {
	sum += boots[ind];
      }
      auto avg = sum / (float) excludes[k].size();
      
      // Now compute the distribution of values WITH this element...
      vector<TYPE> distrib;
      auto excIndex = 0;
      for (auto i = 0; i < nBootstraps; i++) {
	if ((excIndex < excludes[k].size()) && (i == excludes[k][excIndex])) {
	  excIndex++;
	} else {
	  distrib.push_back (boots[i]);
	}
      }

      auto const sz = distrib.size();
      auto leftInterval  = floor (significanceLevel / 2.0 * sz);
      auto rightInterval  = ceil ((1.0 - significanceLevel / 2.0) * sz - 1);

      cout << "avg = " << avg << ", interval = [" << distrib[leftInterval] << "," << distrib[rightInterval] << "]" << endl;

      if ((avg < distrib[leftInterval]) || (avg > distrib[rightInterval])) {
	significant[k] = true;
      } else {
	significant[k] = false;
      }
    }
  }

}

#endif
