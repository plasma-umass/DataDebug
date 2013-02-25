// -*- C++ -*-

#ifndef STATS_H_
#define STATS_H_

#include <assert.h>
#include <math.h>

#include <vector>
using namespace std;

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
  bool mannWhitney (vector<TYPE>& a,
		    vector<TYPE>& b,
		    float significanceLevel = 0.05)
  {
    assert (a.size() == b.size());

    sort (a.begin(), a.end());
    sort (b.begin(), b.end());

    double R[2] = {0.0, 0.0};

    // Compute ranks.
    auto aIndex = 0;
    auto bIndex = 0;
    for (auto i = 0; i < a.size() + b.size(); i++) {
      if (aIndex >= a.size()) {
	R[1] += i + 1;
	bIndex++;
      } else if (bIndex >= b.size()) {
	R[0] += i + 1;
	aIndex++;
      } else {
	if (a[aIndex] < b[bIndex]) {
	  R[0] += i + 1;
	  aIndex++;
	} else if (a[aIndex] > b[bIndex]) {
	  R[1] += i + 1;
	  bIndex++;
	}
	else {
	  // Equal: split the tie.
	  assert (a[aIndex] == b[bIndex]);
	  R[0] += (i + 1)/2.0;
	  R[1] += (i + 1)/2.0;
	  aIndex++;
	  bIndex++;
	}
      }
    }

    double U[2] = { 0.0, 0.0 };
    auto n1 = a.size();
    auto n2 = b.size();
    auto N  = n1 + n2;
    U[0] = R[0] - (n1*(n1+1))/2.0;
    U[1] = R[1] - (n2*(n2+1))/2.0;
    assert (U[0] + U[1] == n1 * n2);
    //    auto mean = (n1 * n2) / 2.0;

    auto u_val = n1 * n2 + n1 * (n1+1)/2 - R[0];
    double mean = n1 * n2 / 2.0;

    //    auto mean = (n1 * n2) / 2.0;
    // without ties, use this:
    auto stddev = sqrt((n1 * n2 * (N + 1.0))/12.0);


    cout << "min U = " << min(R[0],R[1]) << endl;
    cout << "uval = " << u_val << endl;
    cout << "mean = " << mean << endl;
    cout << "alternate z = " << (u_val - mean) / stddev << endl;

    // Since we can have ties, we should use an adjustment. This would
    // entail counting the number of tied ranks, which is reasonably
    // complex. We punt this for now.
    float zScore = fabs(min(U[0],U[1]) - mean) / stddev;
    cout << zScore << endl;
 
    // For now we hard code this. FIX ME.
    if (zScore > 6.0) {
      return true;
    } else {
      return false;
    }
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


}

#endif
