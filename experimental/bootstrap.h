// -*- C++ -*-

#ifndef BOOTSTRAP_H_
#define BOOTSTRAP_H_

#include <stdlib.h>
#include <vector>

using namespace std;

namespace bootstrap {
  
  /// @brief Generate a bootstrapped sample from the input distribution.
  template <class TYPE>
  void complete (const vector<TYPE>& in,
		 vector<TYPE>& out)
  {
    assert (in.size() == out.size());
    const auto N = in.size();
    for (auto& x : out) {
      x = in[lrand48() % N];
    }
  }

  /// @brief Generate a bootstrapped sample from the input distribution.
  template <class TYPE>
  void completeTracked (const vector<TYPE>& in,
			vector<TYPE>& out,
			vector<bool>& included)
  {
    assert (in.size() == out.size());
    fill (included.begin(), included.end(), false);
    const auto N = in.size();
    for (auto& x : out) {
      auto index = lrand48() % N;
      included[index] = true;
      x = in[index];
    }
  }

  /// @brief Generate a bootstrapped sample from the input distribution,
  /// excluding one element.
  template <class TYPE>
  void exclusiveTracked (unsigned long excludeIndex,
			 const vector<TYPE>& in,
			 vector<TYPE>& out,
			 vector<bool>& included)
  {
    assert (in.size() == out.size());
    const auto N = in.size();
    fill (included.begin(), included.end(), false);
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
      included[index] = true;
    }
  }

  /// @brief Generate a bootstrapped sample from the input distribution,
  /// excluding one element.
  template <class TYPE>
  void exclusive (unsigned long excludeIndex,
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
    }
  }
  
}

#endif
