// -*- C++ -*-

#ifndef STATS_H_
#define STATS_H_

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

}

#endif
