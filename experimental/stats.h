// -*- C++ -*-

#ifndef STATS_H_
#define STATS_H_

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
  bool meanDistance (vector<TYPE>& a,
		     vector<TYPE>& b,
		     float significanceLevel = 0.05)
  {
    assert (a.size() == b.size());

    sort (a.begin(), a.end());
    sort (b.begin(), b.end());

    auto aAverage = average (a);
    auto bAverage = average (b);

    auto leftInterval  = floor(significanceLevel / 2.0 * a.size());
    auto rightInterval = ceil((1.0 - significanceLevel / 2.0) * a.size());

    bool result;
    if ((aAverage < b[leftInterval]) ||
	(aAverage > b[rightInterval]) ||
	(bAverage < a[leftInterval]) ||
	(bAverage > a[rightInterval])) {
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
